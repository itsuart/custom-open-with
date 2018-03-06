#include <Windows.h>
#include <GuidDef.h>
#include <Shobjidl.h>
#include <Shlobj.h>

#include <vector>
#include <string>
#include <array>
#include <memory>

using std::vector;
using std::wstring;
using std::array;

static const wchar_t* EXTENSION_GUID_TEXT{ L"{7BA11196-950C-4CC8-81E8-9853F514127F}" };
static constexpr GUID EXTENSION_GUID = { 0x7ba11196, 0x950c, 0x4cc8,{ 0x81, 0xe8, 0x98, 0x53, 0xf5, 0x14, 0x12, 0x7f } };

static constexpr int MAX_WIDE_PATH_LENGTH = 32767;
static const std::wstring UNC_PATH_PREFIX(LR"(\\)");
static const std::wstring WIDE_PATH_PREFIX(LR"(\\?\)");

static std::unique_ptr<wchar_t, decltype(CoTaskMemFree)*> GetUserDocumentsFolderPath() {
    wchar_t* pMyDocuments = nullptr;
    SHGetKnownFolderPath(FOLDERID_Documents, KF_FLAG_CREATE, 0, &pMyDocuments);
    return std::unique_ptr<wchar_t, decltype(CoTaskMemFree)*>(pMyDocuments, CoTaskMemFree);
}


enum class Handlers : unsigned {
    None = 0, // is this even possible?
    Everything = 1,         // L"\\Everything"
    Folders = 2,            // L"\\Folders"
    ExtensionlessFiles = 4, // L"\\Extensionless Files"
    SpecificExtension = 8,  // L"\\Files by Extension"
    AllFiles = 16           // L"\\All files"
};

Handlers operator | (Handlers a, Handlers b) {
    return static_cast<Handlers>(static_cast<unsigned>(a) | static_cast<unsigned>(b));
}

class MyExtension : public IUnknown, IContextMenu, IShellExtInit {
private :
    long m_nRefs = 1;
    
    vector<wstring> m_itemPaths;
    
    wstring m_handlersRoot;

    vector<wstring> m_handlers;

    bool GetFileExtension(const wstring& path, wstring& extension) const {
        for (auto current = path.rbegin(), end = path.rend(); current != end; ++current) {
            if (*current == L'\\' || *current == L'/') {
                // we reached path separator - there is no extension
                return false;
            } else if (*current == L'.') {
                auto pos = end - current - 1; //I want that dot too
                extension = path.substr(pos);
                return true;
            }
        }
        return false;
    }

    Handlers DecideHandlers(wstring& commonExtensionIfAny) const {
        bool haveExtensionlessFiles = false;
        bool haveFilesWithExtension = false;
        bool haveFolders = false;
        bool haveFiles = false;
        bool haveDifferentExtensions = false;
        wstring commonExtension;
        for (const auto& aPathToThing : m_itemPaths) {
            bool isDirectory = FILE_ATTRIBUTE_DIRECTORY == (FILE_ATTRIBUTE_DIRECTORY & GetFileAttributesW(aPathToThing.c_str()));
            if (isDirectory) {
                haveFolders = true;
            } else {
                haveFiles = true;

                // ok what kind of file are you? do you have an extension?
                wstring extension;
                bool haveExtension = GetFileExtension(aPathToThing, extension);
                if (haveExtension) {
                    haveFilesWithExtension = true;

                    if (commonExtension.empty()) {
                        commonExtension = extension;
                    } else {
                        if (haveDifferentExtensions) {
                            //no need to compare anything
                            continue;
                        }

                        if (CSTR_EQUAL != CompareStringOrdinal(commonExtension.c_str(), commonExtension.length(), extension.c_str(), extension.length(), true)) {
                            //so no, new extension is not the same we saw before
                            haveDifferentExtensions = true;
                        }
                    }
                } else {
                    haveExtensionlessFiles = true;
                }
            }
        }

        // decision time
        
        if ( (! haveFiles) && (! haveFolders)) {
            // possible if there is nothing selected that on filesystem
            return Handlers::None;
        }
        
        if (haveFiles && haveFolders) {
            //only Everything is eligible
            return Handlers::Everything;
        }

        if (haveFiles && (!haveFolders)) {
            Handlers result = Handlers::Everything | Handlers::AllFiles;
            //only files, eh? what kind?
            if (haveExtensionlessFiles && haveFilesWithExtension) {
                // we can't specify handlers any further
                return result;
            }

            if (haveExtensionlessFiles && (!haveFilesWithExtension)) {
                return result | Handlers::ExtensionlessFiles;
            }

            if (haveFilesWithExtension && (!haveExtensionlessFiles)) {
                // maybe they have common extension?
                if (haveDifferentExtensions) {
                    // some extensions are different - no special case for that
                    return result;
                } else {
                    // all files has same extension, hurray!
                    commonExtensionIfAny = commonExtension;
                    return result | Handlers::SpecificExtension;
                }
            }

            // a crazy case, not sure if it even possible
            if ((!haveFilesWithExtension) && (!haveExtensionlessFiles)) {
                OutputDebugStringW(L"--------------- CRAZY THING HAPPENED: (!haveFilesWithExtension) && (!haveExtensionlessFiles)");
                return result;
            }
        }

        if (haveFolders && (!haveFiles)) {
            return Handlers::Everything | Handlers::Folders;
        }

        // shouldn't be possible
        return Handlers::None;
    }

public :
    static long m_nInstances;

    MyExtension() {
        auto myDocuments = GetUserDocumentsFolderPath();
        m_handlersRoot = myDocuments.get();

        if (m_handlersRoot.substr(0, UNC_PATH_PREFIX.length()) != UNC_PATH_PREFIX) {
            m_handlersRoot.insert(0, WIDE_PATH_PREFIX);
        }

        m_handlersRoot.append(L"\\Open With Handlers for");

        InterlockedIncrement(&m_nInstances);
    }

    virtual ~MyExtension() {
        InterlockedDecrement(&m_nInstances);
    }

    virtual ULONG STDMETHODCALLTYPE AddRef() override {
        return InterlockedIncrement(&m_nRefs);
    }

    virtual ULONG STDMETHODCALLTYPE Release() override {
        auto nRefs = InterlockedDecrement(&m_nRefs);
        if (nRefs == 0) {
            delete this;
        }
        return nRefs;
    }

    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID requestedIID, void** ppv) override {
        if (IsEqualGUID(requestedIID, EXTENSION_GUID)) {
            *ppv = this;
            AddRef();
            return S_OK;
        }
        if (IsEqualGUID(requestedIID, IID_IUnknown)) {
            *ppv = static_cast<IUnknown*>(this);
            AddRef();
            return S_OK;
        }

        if (IsEqualGUID(requestedIID, IID_IContextMenu)) {
            *ppv = static_cast<IContextMenu*>(this);
            AddRef();
            return S_OK;
        }

        if (IsEqualGUID(requestedIID, IID_IShellExtInit)) {
            *ppv = static_cast<IShellExtInit*>(this);
            AddRef();
            return S_OK;
        }
       
        return E_NOINTERFACE;
    }

    virtual HRESULT STDMETHODCALLTYPE Initialize(LPCITEMIDLIST pIDFolder, IDataObject* pDataObj, HKEY hRegKey) override {
        UNREFERENCED_PARAMETER(pIDFolder);
        UNREFERENCED_PARAMETER(hRegKey);

        m_itemPaths.clear();

        FORMATETC   fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        STGMEDIUM   medium;
        if (SUCCEEDED(pDataObj->GetData(&fe, &medium))) {
            // Get the count of files dropped.
            auto uCount = DragQueryFile((HDROP)medium.hGlobal, (UINT)-1, NULL, 0);

            for (UINT i = 0; i < uCount; i += 1) {
                array<wchar_t, MAX_WIDE_PATH_LENGTH> buffer{ 0 };
                auto nCharsCopied = DragQueryFileW((HDROP)medium.hGlobal, i, buffer.data(), buffer.size());
                if (nCharsCopied) {
                    m_itemPaths.emplace_back(buffer.data());
                } else {
                    OutputDebugStringA("My Open With Extension: Failed to get name of selected item");
                }

            }

            ReleaseStgMedium(&medium);
        }
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) override {

        if (CMF_DEFAULTONLY == (CMF_DEFAULTONLY & uFlags)
            || CMF_VERBSONLY == (CMF_VERBSONLY & uFlags)
            || CMF_NOVERBS == (CMF_NOVERBS & uFlags)) {

            return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
        }

        //OK let as see what handlers we are looking for, starting from most specific

        wstring commonExtension;
        Handlers handlers = DecideHandlers(commonExtension);

        HMENU handlersMenu = CreateMenu();
        
        // order is from top to bottom: most specialized -> least specialized, so
        // SpecializedExtenison
        // <separator if ^ not empty>
        // AllFiles (or ExtensionlessFiles, or Folders)
        // <separator if ^ not empty>
        // AllFiles (if prev. was Extensionless)
        // <separator if ^ not empty>
        // Everything 
        
        // <separator if ^ not empty>
        // "Open handlers folder"
        InsertMenuW(handlersMenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + 1, L"Open handlers folder");

        //in highly unlikely case where is no room left in the menu:
        if (idCmdFirst + 2 > idCmdLast) {
            //we don't add any menu enries of ours
            m_itemPaths.clear();
            return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
        }

        //insert our menu items: separator, copy and move in that order.
        InsertMenu(hmenu, -1, MF_BYPOSITION | MF_SEPARATOR, idCmdFirst + 0, NULL);
        InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING | MF_POPUP, (UINT_PTR)handlersMenu, L"My Open with");
     
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 1 + 1);
    }

    virtual HRESULT STDMETHODCALLTYPE GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT* pwReserved, LPSTR pszName, UINT cchMax) override {
        UNREFERENCED_PARAMETER(idCmd);
        UNREFERENCED_PARAMETER(pwReserved);
        UNREFERENCED_PARAMETER(pszName);
        UNREFERENCED_PARAMETER(cchMax);

        if ((uFlags == GCS_VERBW) || (uFlags == GCS_VALIDATEW)) {
            return S_OK;
        } else {
            return S_FALSE;
        }

        return E_INVALIDARG;
    }

    virtual HRESULT STDMETHODCALLTYPE InvokeCommand(LPCMINVOKECOMMANDINFO pCommandInfo) override {
        const auto pVerb = pCommandInfo->lpVerb;

        if (HIWORD(pVerb)) {
            return E_INVALIDARG;
        }

        const WORD itemIndex = (WORD)pVerb;
        if (itemIndex == m_handlers.size() + 1) {
            //one off - 'open handlers directory'
            ShellExecuteW(NULL, L"explore", m_handlersRoot.c_str(), NULL, NULL, SW_SHOW);
        }
        
        return S_OK;
    }
};

long MyExtension::m_nInstances = 0;

class MyClassFactory : public IClassFactory {
private:
    long m_nRefs = 1;

public:
    static long m_nInstances;
    static long m_nLocks;

    MyClassFactory() {
        InterlockedIncrement(&m_nInstances);
    }

    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID requestedIID, void** ppv) override {
        // Check if the GUID matches an IClassFactory or IUnknown GUID.
        if (!IsEqualGUID(requestedIID, IID_IUnknown) &&
            !IsEqualGUID(requestedIID, IID_IClassFactory)) {
            // It isn't. Clear his handle, and return E_NOINTERFACE.
            *ppv = 0;
            return E_NOINTERFACE;
        }

        *ppv = this;
        AddRef();
        return S_OK;
    }

    virtual ULONG STDMETHODCALLTYPE AddRef() override {
        return InterlockedIncrement(&m_nRefs);
    }

    virtual ULONG STDMETHODCALLTYPE Release() override {
        auto nRefs = InterlockedDecrement(&m_nRefs);
        if (nRefs == 0) {
            delete this;
            InterlockedDecrement(&m_nInstances);
        }
        return nRefs;
    }

    virtual HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *pUnkOuter, REFIID pRequestedIID, void **ppvObject) override {
        if (ppvObject == NULL) {
            return E_POINTER;
        }

        // Assume an error by clearing caller's handle.
        *ppvObject = 0;

        // We don't support aggregation in IExample.
        if (pUnkOuter) {
            return CLASS_E_NOAGGREGATION;
        }

        if (IsEqualGUID(pRequestedIID, IID_IUnknown)
            || IsEqualGUID(pRequestedIID, IID_IContextMenu)
            || IsEqualGUID(pRequestedIID, IID_IShellExtInit)
            ) {

            MyExtension* pExt = new MyExtension();
            return pExt->QueryInterface(pRequestedIID, ppvObject);
        }

        return E_NOINTERFACE;
    }

    virtual HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) override {
        if (fLock) {
            InterlockedIncrement(&m_nLocks);
        } else {
            InterlockedDecrement(&m_nLocks);
        }

        return S_OK;
    };
};

long MyClassFactory::m_nLocks = 0;
long MyClassFactory::m_nInstances = 0;


/*
COM subsystem will call this function in attempt to retrieve an implementation of an IID.
Usually it's IClassFactory that is requested (but could be anything else?)
*/
HRESULT __stdcall DllGetClassObject(REFCLSID pCLSID, REFIID pIID, void** ppv) {
    if (ppv == NULL) {
        return E_POINTER;
    }

    *ppv = NULL;
    if (!IsEqualGUID(pCLSID, EXTENSION_GUID)) {
        //that's not our CLSID, ignore
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    if (!IsEqualGUID(pIID, IID_IClassFactory)) {
        //we only providing ClassFactories in this function
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    //alright, return our implementation of IClassFactory
    *ppv = new MyClassFactory;
    return S_OK;
}

HRESULT PASCAL DllCanUnloadNow() {
    if (   MyClassFactory::m_nLocks == 0 
        && MyClassFactory::m_nInstances == 0 
        && MyExtension::m_nInstances == 0) {
        return S_OK;
    }
    return S_FALSE;
}


BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

