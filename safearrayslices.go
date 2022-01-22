// +build windows

package ole

import (
	"unsafe"
)

func safeArrayFromByteSlice(slice []byte) *SafeArray {
	array, _ := safeArrayCreateVector(VT_UI1, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []byte to SAFEARRAY")
	}

	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
	}
	return array
}

func safeArrayFromStringSlice(slice []string) *SafeArray {
	array, _ := safeArrayCreateVector(VT_BSTR, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []string to SAFEARRAY")
	}
	// SysAllocStringLen(s)
	for i, v := range slice {
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(SysAllocStringLen(v))))
	}
	return array
}

func safeArrayFromIDispatchSlice(slice []*IDispatch) *SafeArray {
	array, _ := safeArrayCreateVector(VT_DISPATCH, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []*IDispatch to SAFEARRAY")
	}

	for i, v := range slice {
		// note: v not &v as this is array of pointers!
		safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(v)))
	}
	return array
}
