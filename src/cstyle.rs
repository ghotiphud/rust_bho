use winapi;
use winapi::{REFIID, HRESULT, ULONG, BOOL, c_void,
    NOERROR, E_NOINTERFACE, CLASS_E_NOAGGREGATION,
    IUnknown, E_FAIL};
use user32;
use super::*;
use std::ptr;
use std::sync::atomic::{AtomicUsize, Ordering};

use std::ffi::OsStr;
use std::io::Error;
use std::iter::once;
use std::os::windows::ffi::OsStrExt;

fn debug_message_box(msg: &str) {
    let wide: Vec<u16> = OsStr::new(msg).encode_wide().chain(once(0)).collect();
    let ret = unsafe {
        user32::MessageBoxW(ptr::null_mut(), wide.as_ptr(), wide.as_ptr(), winapi::MB_OK)
    };

    if ret == 0 {
        println!("Failed: {:?}", Error::last_os_error());
    }
}

//       IEExtension    IEExtensionVtbl
//       +--------+     +--------------+
// p --> | lpVtbl | --> |QueryInterface|
//       |        |     |AddRef        |
//       | data   |     |Release       |
//       |        |     |SetSite       |
//       +--------+     |GetSite       |
//                      |              |
//                      +--------------+


// IUnknown
// -----------
// HRESULT QueryInterface(
//   [in]  REFIID riid,
//   [out] void   **ppvObject
// );
//
// ULONG AddRef();
//
// ULONG Release();


// IObjectWithSite
// -----------------
// HRESULT SetSite(
//   [in] IUnknown *pUnkSite
// );
//
// HRESULT GetSite(
//   [in]  REFIID riid,
//   [out] void   **ppvSite
// );

#[repr(C)]
#[derive(Debug)]
pub struct IEExtension {
    lpVtbl: *const IEExtensionVtbl,
    counter: ULONG,
    pub site: *mut IUnknown,
}

impl IEExtension {
    pub fn new(counter: ULONG) -> Self {
        IEExtension { lpVtbl: IEExtensionVtbl_STATIC, counter: counter, site: ptr::null_mut() }
    }

    #[inline]
    pub unsafe fn SetSite(&mut self, pUnkSite: *mut IUnknown) -> HRESULT {
        ((*self.lpVtbl).SetSite)(self, pUnkSite)
    }
    #[inline]
    pub unsafe fn GetSite(&mut self, riid: REFIID, ppvSite: *mut *mut c_void)
     -> HRESULT {
        ((*self.lpVtbl).GetSite)(self, riid, ppvSite)
    }
}

impl ::std::ops::Deref for IEExtension {
    type Target = IUnknown;
    #[inline]
    fn deref(&self) -> &IUnknown {
        unsafe { ::std::mem::transmute(self) }
    }
}
impl ::std::ops::DerefMut for IEExtension {
    #[inline]
    fn deref_mut(&mut self) -> &mut IUnknown {
        unsafe { ::std::mem::transmute(self) }
    }
}

#[repr(C)]
struct IEExtensionVtbl {
    pub QueryInterface: unsafe extern "system" 
        fn(This: *mut IEExtension, riid: REFIID, ppvObject: *mut *mut c_void) -> HRESULT,
    pub AddRef: unsafe extern "system" 
        fn(This: *mut IEExtension) -> ULONG,
    pub Release: unsafe extern "system" 
        fn(This: *mut IEExtension) -> ULONG,
    pub SetSite: unsafe extern "system" 
        fn(This: *mut IEExtension, pUnkSite: *mut IUnknown) -> HRESULT,
    pub GetSite: unsafe extern "system" 
        fn(This: *mut IEExtension, riid: REFIID, ppvSite: *mut *mut c_void) -> HRESULT,
}

const IEExtensionVtbl_STATIC: &'static IEExtensionVtbl = 
&IEExtensionVtbl{
    // IUnknownVtbl
    QueryInterface: {
        unsafe extern "system" fn IMPL(This: *mut IEExtension, riid: REFIID, ppvObject: *mut *mut c_void) -> HRESULT  { 
            // Check if GUID matches 
            if IsEqualGUID(riid, &IID_IEExtension) != 0 &&
                IsEqualGUID(riid, &IID_IObjectWithSite) != 0 &&
                IsEqualGUID(riid, &IID_IUnknown) != 0 {
                // Unrecognized GUID, clear handle and return E_NOINTERFACE
                *ppvObject = ptr::null_mut();
                return E_NOINTERFACE;
            }

            // GUID matches
            // 1. Set pointer to this object
            *ppvObject = This as (*mut c_void);
            // 2. Reference Count
            ((*(*This).lpVtbl).AddRef)(This);
            // 3. Return NOERROR
            NOERROR
        }
        IMPL
    },
    AddRef: {
        unsafe extern "system" fn IMPL(This: *mut IEExtension) -> ULONG {
            // println!("This AddRef: {:?}", *This);
            (*This).counter += 1;
            (*This).counter
        }
        IMPL
    },
    Release: {
        unsafe extern "system" fn IMPL(This: *mut IEExtension) -> ULONG {
            // println!("This Release: {:?}", *This);
            let count = {
                (*This).counter -= 1;
                (*This).counter
            };

            if count == 0 {
                Box::from_raw(This);
                ::OUTSTANDING_OBJECTS.fetch_sub(1, Ordering::SeqCst);
            }

            count
        }
        IMPL
    },

    // IObjectWithSiteVtbl
    SetSite: {
        unsafe extern "system" fn IMPL(This: *mut IEExtension, pUnkSite: *mut IUnknown) -> HRESULT {
            println!("extn SetSite");

            // AddRef so object doesn't die while we're working with it.
            //if(pUnkSite != ptr::null_mut()) { pUnkSite.AddRef(); }

            (*This).site = pUnkSite;

            debug_message_box("SetSite");

            NOERROR
        }
        IMPL
    },
    GetSite: {
        unsafe extern "system" fn IMPL(This: *mut IEExtension, riid: REFIID, ppvSite: *mut *mut c_void) -> HRESULT {
            println!("extn GetSite");
            // TODO: implement
            *ppvSite = ptr::null_mut();

            if (*This).site == ptr::null_mut() { return E_FAIL; }

            debug_message_box("GetSite");

            (*This).QueryInterface(riid, ppvSite)
        }
        IMPL
    },
};

// IClassFactory
// --------------
// HRESULT CreateInstance(
//   [in]  IUnknown *pUnkOuter,
//   [in]  REFIID   riid,
//   [out] void     **ppvObject
// );
//
// HRESULT LockServer(
//   [in] BOOL fLock
// );

#[repr(C)]
pub struct IEExtensionClassFactory {
    lpVtbl: *const IEExtensionClassFactoryVtbl,
}

impl IEExtensionClassFactory {
    #[inline]
    pub unsafe fn QueryInterface(&self, riid: REFIID, ppvObject: *mut *mut c_void) -> HRESULT {
        ((*self.lpVtbl).QueryInterface)(self, riid, ppvObject)
    }
    #[inline]
    pub unsafe fn CreateInstance(&self, pUnkOuter: *mut IUnknown, riid: REFIID, ppvObject: *mut *mut c_void) -> HRESULT {
        ((*self.lpVtbl).CreateInstance)(self, pUnkOuter, riid, ppvObject)
    }
}

#[repr(C)]
struct IEExtensionClassFactoryVtbl {
    pub QueryInterface: unsafe extern "system" 
        fn(This: *const IEExtensionClassFactory, riid: REFIID, ppvObject: *mut *mut c_void) -> HRESULT,
    pub AddRef: unsafe extern "system" 
        fn(This: *const IEExtensionClassFactory) -> ULONG,
    pub Release: unsafe extern "system" 
        fn(This: *const IEExtensionClassFactory) -> ULONG,
    pub CreateInstance: unsafe extern "system"
        fn(This: *const IEExtensionClassFactory, pUnkOuter: *mut IUnknown, riid: REFIID, ppvObject: *mut *mut c_void) -> HRESULT,
    pub LockServer: unsafe extern "system"
        fn(This: *const IEExtensionClassFactory, fLock: BOOL) -> HRESULT,
}

pub const IEExtensionClassFactory_STATIC: &'static IEExtensionClassFactory = &IEExtensionClassFactory { 
    lpVtbl: IEExtensionClassFactoryVtbl_STATIC,
};

const IEExtensionClassFactoryVtbl_STATIC: &'static IEExtensionClassFactoryVtbl = 
&IEExtensionClassFactoryVtbl {
    // IUnknownVtbl
    QueryInterface: {
        unsafe extern "system" fn IMPL(This: *const IEExtensionClassFactory, riid: REFIID, ppvObject: *mut *mut c_void) -> HRESULT  { 
            // Check if GUID matches 
            if IsEqualGUID(riid, &IID_IClassFactory) != 0 &&
                IsEqualGUID(riid, &IID_IUnknown) != 0 {
                // Unrecognized GUID, clear handle and return E_NOINTERFACE
                *ppvObject = ptr::null_mut();
                return E_NOINTERFACE;
            }

            // GUID matches
            // 1. Set pointer to this object
            *ppvObject = This as (*mut c_void);
            // 2. Reference Count not needed, static
            ((*(*This).lpVtbl).AddRef)(This); 
            // 3. Return NOERROR
            NOERROR
        }
        IMPL
    },
    AddRef: {
        unsafe extern "system" fn IMPL(This: *const IEExtensionClassFactory) -> ULONG {
            // Static, no ref counting
            1
        }
        IMPL
    },
    Release: {
        unsafe extern "system" fn IMPL(This: *const IEExtensionClassFactory) -> ULONG {
            // Static, no ref counting
            1
        }
        IMPL
    },

    // IClassFactory
    CreateInstance: {
        unsafe extern "system" fn IMPL(This: *const IEExtensionClassFactory, 
                                        pUnkOuter: *mut IUnknown, 
                                        riid: REFIID, 
                                        ppvObject: *mut *mut c_void) -> HRESULT {
            println!("cls factory CreateInstance");
            let hr: HRESULT;
            *ppvObject = ptr::null_mut();

            if pUnkOuter != ptr::null_mut() {
                hr = CLASS_E_NOAGGREGATION;
            } else {
                // create a new instance of IEExtension on the heap
                let mut ie_ext_ptr = Box::new(IEExtension::new(1));
                // turn it into a raw pointer to allow manual memory management
                let mut ie_ext_ptr = Box::into_raw(ie_ext_ptr);

                hr = ((*(*ie_ext_ptr).lpVtbl).QueryInterface)(ie_ext_ptr, riid, ppvObject);

                ((*(*ie_ext_ptr).lpVtbl).Release)(ie_ext_ptr);

                // object allocated, ppvObject takes ownership
                if hr == NOERROR {
                    ::OUTSTANDING_OBJECTS.fetch_add(1, Ordering::SeqCst);
                }
            }

            hr
        }
        IMPL
    },
    LockServer: {
        unsafe extern "system" fn IMPL(This: *const IEExtensionClassFactory, fLock: BOOL) -> HRESULT {
            println!("cls factory LockServer");
            if fLock != 0 {
                ::LOCK_COUNT.fetch_add(1, Ordering::SeqCst);
            }else{
                ::LOCK_COUNT.fetch_sub(1, Ordering::SeqCst);
            }

            NOERROR
        }
        IMPL
    },
};