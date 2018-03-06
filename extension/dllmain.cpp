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

class MyExtension : public IUnknown, IContextMenu, IShellExtInit {
private :
    long m_nRefs = 1;
    vector<wstring> m_itemPaths;
    wstring m_handlersRoot;
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

        



        //in highly unlikely case where is no room left in the menu:
        if (idCmdFirst + 2 > idCmdLast) {
            //we don't add any menu enries of ours
            m_itemPaths.clear();
            return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
        }

        //insert our menu items: separator, copy and move in that order.
        InsertMenu(hmenu, -1, MF_BYPOSITION | MF_SEPARATOR, idCmdFirst + 0, NULL);
        InsertMenu(hmenu, -1, MF_BYPOSITION | MF_STRING, idCmdFirst + 1, L"My Open with");
     
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 1 + 1); //shouldn't this be +idCmdFirst?
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

