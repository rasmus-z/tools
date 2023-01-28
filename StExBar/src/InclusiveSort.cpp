#include "stdafx.h"
#include "SRBand.h"
#include "StringUtils.h"
#include "OnOutOfScope.h"
#include "Pidl.h"
#include <regex>
#include <propkey.h> //needed for PROPERTYKEY's in Column lists
#include "Filter.h"
#include "IListView.h"

bool CDeskBand::InclusiveSort()
{
    bool              bReturn = false;
    IServiceProvider* pServiceProvider;
    if (SUCCEEDED(m_pSite->QueryInterface(IID_IServiceProvider, (LPVOID*)&pServiceProvider)))
    {
        IShellBrowser* pShellBrowser;
        if (SUCCEEDED(pServiceProvider->QueryService(SID_SShellBrowser, IID_IShellBrowser, (LPVOID*)&pShellBrowser)))
        {
            IShellView* pShellView;
            if (SUCCEEDED(pShellBrowser->QueryActiveShellView(&pShellView)))
            {
                IFolderView* pFolderView;
                if (SUCCEEDED(pShellView->QueryInterface(IID_IFolderView, (LPVOID*)&pFolderView)))
                {
                    // hooray! we got the IFolderView interface!
                    // that means the explorer is active and well :)
                    IShellFolderView* pShellFolderView;
                    if (SUCCEEDED(pShellView->QueryInterface(IID_IShellFolderView, (LPVOID*)&pShellFolderView)))
                    {
                        // the first thing we do is to deselect all already selected entries
                        pFolderView->SelectItem(NULL, SVSI_DESELECTOTHERS);

                        // but we also need the IShellFolder interface because
                        // we need its GetDisplayNameOf() method
                        IPersistFolder2* pPersistFolder;
                        if (SUCCEEDED(pFolderView->GetFolder(IID_IPersistFolder2, (LPVOID*)&pPersistFolder)))
                        {
                            LPITEMIDLIST curFolder;
                            pPersistFolder->GetCurFolder(&curFolder);
                            if (ILIsEqual(m_currentFolder, curFolder))
                            {
                                CoTaskMemFree(curFolder);
                            }
                            else
                            {
                                CoTaskMemFree(m_currentFolder);
                                m_currentFolder = curFolder;
                                for (size_t i = 0; i < m_noShows.size(); ++i)
                                {
                                    CoTaskMemFree(m_noShows[i]);
                                }
                                m_noShows.clear();
                            }
                            IShellFolder* pShellFolder;
                            if (SUCCEEDED(pPersistFolder->QueryInterface(IID_IShellFolder, (LPVOID*)&pShellFolder)))
                            {
                                // our next task is to enumerate all the
                                // items in the folder view and select those
                                // which match the text in the edit control

                                int nCount = 0;
                                if (SUCCEEDED(pFolderView->ItemCount(SVGIO_ALLVIEW, &nCount)))
                                {
                                    pShellFolderView->SetRedraw(FALSE);
                                    HWND    listView = GetListView32(pShellView);
                                    LRESULT viewType = 0;
                                    if (listView)
                                    {
                                        // inserting items in the list view if the list view is set to
                                        // e.g., LV_VIEW_LIST is painfully slow. So save the current view
                                        // and set it to LV_VIEW_DETAILS (which is much faster for inserting)
                                        // and restore the view after we're done.
                                        viewType = SendMessage(listView, LVM_GETVIEW, 0, 0);
                                        SendMessage(listView, LVM_SETVIEW, LV_VIEW_DETAILS, 0);

                                        IListView_Win7* pLvw = NULL;
                                        SendMessage(listView, LVM_QUERYINTERFACE, reinterpret_cast<WPARAM>(&IID_IListView_Win7), reinterpret_cast<LPARAM>(&pLvw));
                                        if (pLvw)
                                        {
                                            pLvw->EnableAlphaShadow(true);
                                            //pLvw->
                                            pLvw->Release();
                                        }
                                    }
                                    std::vector<LPITEMIDLIST> noShows;
                                    noShows.reserve(nCount);
                                    for (int i = 0; i < nCount; ++i)
                                    {
                                        LPITEMIDLIST pidl;
                                        if (SUCCEEDED(pFolderView->Item(i, &pidl)))
                                        {
                                            if (false)
                                            {
                                                // remove now shown items which are in the no-show list
                                                // this is necessary since we don't get a notification
                                                // if the shell refreshes its view
                                                for (std::vector<LPITEMIDLIST>::iterator it = m_noShows.begin(); it != m_noShows.end(); ++it)
                                                {
                                                    if (HRESULT_CODE(pShellFolder->CompareIDs(SHCIDS_CANONICALONLY, *it, pidl)) == 0)
                                                    {
                                                        m_noShows.erase(it);
                                                        break;
                                                    }
                                                }
                                                CoTaskMemFree(pidl);
                                            }
                                            else
                                            {
                                                UINT puItem = 0;
                                                // pShellFolderView->UpdateObject
                                                if (pShellFolderView->RemoveObject(pidl, &puItem) == S_OK)
                                                {
                                                    i--;
                                                    nCount--;
                                                    noShows.push_back(pidl);
                                                }
                                            }
                                        }
                                    }

                                    // TODO: Sort noShows
                                    auto g_SortColumn       = 0;
                                    auto g_sortAscending    = false;
                                    auto m_UseInclusiveSort = false;
                                    if (m_UseInclusiveSort)
                                    {
                                        IFolderView2* pFolderView2;
                                        auto          hr = pShellView->QueryInterface(IID_PPV_ARGS(&pFolderView2));
                                        // IFolderView2* pFolderView2;
                                        // if (SUCCEEDED(pShellView->QueryInterface(IID_IFolderView2, (LPVOID*)&pFolderView2)))
                                        if (SUCCEEDED(hr))
                                        {
                                            PROPERTYKEY sortProp;
                                            int         cColumns;
                                            if (SUCCEEDED(pFolderView2->GetSortColumnCount(&cColumns)))
                                            {
                                                std::vector<SORTCOLUMN> columns(cColumns);
                                                hr = pFolderView2->GetSortColumns(&columns[0], cColumns);
                                                if (SUCCEEDED(hr))
                                                {
                                                    // TODO: explorer can be sorted on more than one column using Shift+Click
                                                    for (int i = 0; i < cColumns; i++)
                                                    // for (unsigned i = cColumns; i-- > 0;)
                                                    {
                                                        // Sort direction
                                                        g_sortAscending = columns[0].direction > 0 ? true : false;

                                                        sortProp        = columns[i].propkey;
                                                    }
                                                }
                                                else
                                                {
                                                    //pLogger->error(CStringUtils::Format("Couldn't GetSortColumns: %d", hr));
                                                }
                                            }
                                            else
                                            {
                                                //pLogger->error("Couldn't GetSortColumnCount!");
                                            }

                                            IColumnManager* pcm;
                                            hr = pFolderView2->QueryInterface(&pcm);
                                            if (SUCCEEDED(hr))
                                            {
                                                UINT visibleColumnCount;
                                                if (SUCCEEDED(pcm->GetColumnCount(CM_ENUM_VISIBLE, &visibleColumnCount)))
                                                {
                                                    std::vector<PROPERTYKEY> visibleColumns(visibleColumnCount);
                                                    hr = pcm->GetColumns(CM_ENUM_VISIBLE, &visibleColumns[0], visibleColumnCount);
                                                    for (UINT i = 0; i < visibleColumnCount; i++)
                                                    {
                                                        if (visibleColumns[i] == sortProp)
                                                        {
                                                            g_SortColumn = i;
                                                        }
                                                    }
                                                }
                                                if (true)
                                                {
                                                    CM_COLUMNINFO ci = {sizeof(ci), CM_MASK_STATE};
                                                    hr               = pcm->GetColumnInfo(PKEY_ItemNameDisplay, &ci);
                                                    if (SUCCEEDED(hr))
                                                    {
                                                        ci.dwState = CM_STATE_NOSORTBYFOLDERNESS;
                                                        pcm->SetColumnInfo(PKEY_ItemNameDisplay, &ci);
                                                    }
                                                }
                                                if (true)
                                                {
                                                    CM_COLUMNINFO ci = {sizeof(ci), CM_MASK_STATE};
                                                    hr               = pcm->GetColumnInfo(PKEY_ItemFolderPathDisplay, &ci);
                                                    if (SUCCEEDED(hr))
                                                    {
                                                        ci.dwState = CM_STATE_NOSORTBYFOLDERNESS;
                                                        pcm->SetColumnInfo(PKEY_ItemFolderPathDisplay, &ci);
                                                    }
                                                }
                                                pcm->Release();
                                            }
                                        }
                                        else
                                        {
                                            //pLogger->error(CStringUtils::Format("Couldn't get IFolderView2: %d", hr));
                                        }

                                        auto property             = g_SortColumn;
                                        // TEXT("Usage: %s <file>\n")
                                        // folderpath = CStringUtils::Format(L"%s\\%s (%d)", cwd.c_str(), foldername.c_str(), retrycount);
                                        // auto text = CStringUtils::Format(L"Sort Column set to: %s", g_SortColumn);
                                        // pLogger->info(CStringUtils::Format("Sort Column set to: %d", g_SortColumn));

                                        auto capturedShellFolder  = pShellFolder;
                                        // auto capturedCurFolder   = curFolder;
                                        auto captureSortAscending = g_sortAscending;

                                        auto compare              = [property, captureSortAscending, capturedShellFolder](auto& i1, auto& i2) -> bool {
                                            // Logger::getInstance()->info(CStringUtils::Format("Sort Column capture set to: %d", property));
                                            switch (property)
                                            {
                                                case 0:
                                                    return SortStrings(GetShellItemDisplayName(capturedShellFolder, i1), GetShellItemDisplayName(capturedShellFolder, i2), true);
                                                    break;
                                                    // case 1:
                                                    //     return SortStrings(GetShellItemDisplayName(pShellFolder, i1), GetShellItemDisplayName(pShellFolder, i2), true);
                                                // case 10:
                                                //     GetShellItemFileTime(pShellFolder, curFolder, i1, &f1);
                                                //     GetShellItemFileTime(pShellFolder, curFolder, i2, &f2);
                                                //    return SortHelper::SortFileTime(&f1, &f2, ascending);
                                                case 1:

                                                    // GetShellItemFileTime(pShellFolder, curFolder, i2, &f2);
                                                    // auto p1 = GetShellItemPath(capturedShellFolder, capturedCurFolder, i1);
                                                    // auto p2 = GetShellItemPath(capturedShellFolder, capturedCurFolder, i2);

                                                    auto compare = CompareLastWriteTime(GetShellItemFullyQualifiedName(capturedShellFolder, i1), GetShellItemFullyQualifiedName(capturedShellFolder, i2));

                                                    /*
                                                FILETIME f1;
                                                getFileTime(p1, &f1);
                                                FILETIME f2;
                                                getFileTime(p2, &f2);
                                                auto compare = CompareFileTime(&f1, &f2);
                                                if (compare == -1)
                                                {
                                                    Logger::getInstance()->info(CStringUtils::Format("%s is older than %s", p1, p2));
                                                }
                                                else
                                                {
                                                    Logger::getInstance()->info(CStringUtils::Format("%s is not older than %s, compare is %d", p1, p2, compare));
                                                }
                                                */
                                                    // auto compare = CompareFileTime(f1.ftLastWriteTime, f2.ftLastWriteTime);
                                                    // auto compare = ::_wcsicmp(s2.c_str(), s1.c_str());
                                                    return captureSortAscending ? compare < 0 : compare > 0;
                                                    break;
                                                    // case 2: return sSortHelper::SortStrings(i1->GetEventName(), i2->GetEventName(), si->SortAscending);
                                                    // case 3: return SortHelper::SortNumbers(i1->GetProcessId(), i2->GetProcessId(), si->SortAscending);
                                                    // case 4: return SortHelper::SortStrings(i1->GetProcessName(), i2->GetProcessName(), si->SortAscending);
                                                    // case 5: return SortHelper::SortNumbers(i1->GetThreadId(), i2->GetThreadId(), si->SortAscending);
                                                    // case 6: return SortHelper::SortNumbers(i1->GetEventDescriptor().Opcode, i2->GetEventDescriptor().Opcode, si->SortAscending);

                                                    // default:
                                                    //     return SortStrings(GetShellItemDisplayName(capturedShellFolder, i1), GetShellItemDisplayName(capturedShellFolder, i2), true);
                                                    //     break;
                                            }
                                            return false;
                                        };

                                        // std::map<std::wstring, ULONG> m_selectedItems;    ///< list of items which are selected in the explorer view
                                        sort(noShows.begin(), noShows.end(), compare);
                                    }

                                    // now add all those items again which were removed by a previous filter string
                                    // but don't match this new one
                                    // pShellFolderView->SetObjectCount(5000, SFVSOC_INVALIDATE_ALL|SFVSOC_NOSCROLL);
                                    for (size_t i = 0; i < m_noShows.size(); ++i)
                                    {
                                        LPITEMIDLIST pidlNoShow = m_noShows[i];
                                        // if (CheckDisplayName(pShellFolder, pidlNoShow, filter, bUseRegex))
                                        if (true)
                                        {
                                            m_noShows.erase(m_noShows.begin() + i);
                                            i--;
                                            UINT puItem = (UINT)i;
                                            pShellFolderView->AddObject(pidlNoShow, &puItem);
                                            CoTaskMemFree(pidlNoShow);
                                        }
                                    }
                                    m_noShows.insert(
                                        m_noShows.end(),
                                        std::make_move_iterator(noShows.begin()),
                                        std::make_move_iterator(noShows.end()));
                                    noShows.clear();
                                    if (listView)
                                    {
                                        SendMessage(listView, LVM_SETVIEW, viewType, 0);
                                    }

                                    pShellFolderView->SetRedraw(TRUE);
                                }
                                pShellFolder->Release();
                            }
                            pPersistFolder->Release();
                        }
                        pShellFolderView->Release();
                    }
                    pFolderView->Release();
                }
                pShellView->Release();
            }
            pShellBrowser->Release();
        }
        pServiceProvider->Release();
    }
    return bReturn;
}

std::wstring CDeskBand::GetShellItemDisplayName(IShellFolder* shellFolder, LPITEMIDLIST pidl)
{
    STRRET str;
    if (SUCCEEDED(shellFolder->GetDisplayNameOf(pidl,
                                                // SHGDN_FORPARSING needed to get the extensions even if they're not shown
                                                SHGDN_INFOLDER | SHGDN_FORPARSING,
                                                &str)))
    {
        TCHAR dispname[MAX_PATH];
        StrRetToBuf(&str, pidl, dispname, _countof(dispname));
        try
        {
            std::wstring s = dispname;
            return s;
        }
        catch (std::exception)
        {
        }
    }
    return 0;
}

std::wstring CDeskBand::GetShellItemFullyQualifiedName(IShellFolder* shellFolder, LPITEMIDLIST pidl)
{
    STRRET str;
    if (SUCCEEDED(shellFolder->GetDisplayNameOf(pidl, SHGDN_FORPARSING, &str)))
    {
        TCHAR dispname[MAX_PATH];
        StrRetToBuf(&str, pidl, dispname, _countof(dispname));
        try
        {
            std::wstring s = dispname;
            return s;
        }
        catch (std::exception)
        {
        }
    }
    return 0;
}

std::wstring CDeskBand::GetShellItemPath(IShellFolder* shellFolder, LPITEMIDLIST folderpidl, LPITEMIDLIST pidl)
{
    std::wstring sRet;
    // the pidl we get here is relative!
    ULONG        attribs = SFGAO_FILESYSTEM | SFGAO_FOLDER;
    if (SUCCEEDED(shellFolder->GetAttributesOf(1, (LPCITEMIDLIST*)&pidl, &attribs)))
    {
        wchar_t      buf[MAX_PATH] = {0};

        // create an absolute pidl with the pidl we got above
        LPITEMIDLIST abspidl       = CPidl::Append(folderpidl, pidl);
        if (abspidl)
        {
            if (SHGetPathFromIDList(abspidl, buf))
            {
                try
                {
                    sRet += buf;
                }
                catch (std::exception)
                {
                }
            }
            CoTaskMemFree(abspidl);
        }
    }
    return sRet;
}
