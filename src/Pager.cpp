// Pager.cpp: Superclasses SysPager.

#include "stdafx.h"
#include "Pager.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


Pager::Pager()
{
	pSimpleFrameSite = NULL;

	SIZEL size = {100, 100};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_PGR_MOUSEICON);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// initialize
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP Pager::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IPager, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP Pager::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<Pager>::Load(pPropertyBag, pErrorLog);
	return hr;
}

STDMETHODIMP Pager::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<Pager>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	return hr;
}

STDMETHODIMP Pager::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 11 VT_I4 properties...
	pSize->QuadPart += 11 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 2 VT_I2 properties...
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(SHORT));
	// ...and 6 VT_BOOL properties...
	pSize->QuadPart += 6 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 0 VT_BSTR properties...

	// ...and 1 VT_DISPATCH properties
	pSize->QuadPart += 1 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP Pager::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x12121212/*4x AppID*/) {
		return E_FAIL;
	}
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version > 0x0100) {
		return E_FAIL;
	}

	DWORD atlVersion;
	if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = ReadVariantFromStream;

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG autoScrollFrequency = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I2, &propertyValue))) {
		return hr;
	}
	SHORT autoScrollLinesPerTimeout = propertyValue.iVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I2, &propertyValue))) {
		return hr;
	}
	SHORT autoScrollPixelsPerLine = propertyValue.iVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG borderSize = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL detectDoubleClicks = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	/*if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;*/
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL forwardMouseMessagesToContainedWindow = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OrientationConstants orientation = static_cast<OrientationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL rightToLeftLayout = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ScrollAutomaticallyConstants scrollAutomatically = static_cast<ScrollAutomaticallyConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG scrollButtonSize = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;


	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoScrollFrequency(autoScrollFrequency);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoScrollLinesPerTimeout(autoScrollLinesPerTimeout);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoScrollPixelsPerLine(autoScrollPixelsPerLine);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	// NOTE: Border size is limited by button size.
	hr = put_BorderSize(borderSize);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DetectDoubleClicks(detectDoubleClicks);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	/*hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}*/
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForwardMouseMessagesToContainedWindow(forwardMouseMessagesToContainedWindow);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Orientation(orientation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeftLayout(rightToLeftLayout);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ScrollAutomatically(scrollAutomatically);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ScrollButtonSize(scrollButtonSize);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP Pager::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x12121212/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0100;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.autoScrollFrequency;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I2;
	propertyValue.iVal = properties.autoScrollLinesPerTimeout;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.iVal = properties.autoScrollPixelsPerLine;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderSize;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	/*propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}*/
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.forwardMouseMessagesToContainedWindow);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.orientation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.rightToLeftLayout);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.scrollAutomatically;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.scrollButtonSize;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND Pager::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_PAGESCROLLER_CLASS;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<Pager>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT Pager::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("Pager ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void Pager::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP Pager::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	pSimpleFrameSite = NULL;
	if(pClientSite) {
		pClientSite->QueryInterface(IID_ISimpleFrameSite, reinterpret_cast<LPVOID*>(&pSimpleFrameSite));
	}

	HRESULT hr = CComControl<Pager>::IOleObject_SetClientSite(pClientSite);
	return hr;
}

STDMETHODIMP Pager::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}


//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP Pager::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);

	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}

	if(GetStyle() & PGS_DRAGNDROP) {
		// without the native IDropTarget, auto-scrolling won't work
		CComPtr<IDropTarget> pDropTarget = NULL;
		SendMessage(PGM_GETDROPTARGET, 0, reinterpret_cast<LPARAM>(&pDropTarget));
		if(pDropTarget) {
			DWORD effect = *pEffect;
			pDropTarget->DragEnter(pDataObject, keyState, mousePosition, &effect);
		}
	}
	return S_OK;
}

STDMETHODIMP Pager::DragLeave(void)
{
	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragLeave(&callDropTargetHelper);

	if(GetStyle() & PGS_DRAGNDROP) {
		// without the native IDropTarget, auto-scrolling won't work
		CComPtr<IDropTarget> pDropTarget = NULL;
		SendMessage(PGM_GETDROPTARGET, 0, reinterpret_cast<LPARAM>(&pDropTarget));
		if(pDropTarget) {
			pDropTarget->DragLeave();
		}
	}

	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->DragLeave();
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP Pager::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition, &callDropTargetHelper);

	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}

	if(GetStyle() & PGS_DRAGNDROP) {
		// without the native IDropTarget, auto-scrolling won't work
		CComPtr<IDropTarget> pDropTarget = NULL;
		SendMessage(PGM_GETDROPTARGET, 0, reinterpret_cast<LPARAM>(&pDropTarget));
		if(pDropTarget) {
			DWORD effect = *pEffect;
			pDropTarget->DragOver(keyState, mousePosition, &effect);
		}
	}
	return S_OK;
}

STDMETHODIMP Pager::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);

	if(GetStyle() & PGS_DRAGNDROP) {
		// without the native IDropTarget, auto-scrolling won't work
		CComPtr<IDropTarget> pDropTarget = NULL;
		SendMessage(PGM_GETDROPTARGET, 0, reinterpret_cast<LPARAM>(&pDropTarget));
		if(pDropTarget) {
			DWORD effect = *pEffect;
			pDropTarget->Drop(pDataObject, keyState, mousePosition, &effect);
		}
	}

	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP Pager::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Accessibility:
			*pName = GetResString(IDPC_ACCESSIBILITY).Detach();
			return S_OK;
			break;
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP Pager::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_PGR_APPEARANCE:
		case DISPID_PGR_BORDERSIZE:
		case DISPID_PGR_BORDERSTYLE:
		case DISPID_PGR_MOUSEICON:
		case DISPID_PGR_MOUSEPOINTER:
		case DISPID_PGR_ORIENTATION:
		case DISPID_PGR_SCROLLBUTTONSIZE:
		case DISPID_PGR_SCROLLBUTTONSTATE:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_PGR_AUTOSCROLLFREQUENCY:
		case DISPID_PGR_AUTOSCROLLLINESPERTIMEOUT:
		case DISPID_PGR_AUTOSCROLLPIXELSPERLINE:
		case DISPID_PGR_DETECTDOUBLECLICKS:
		case DISPID_PGR_DISABLEDEVENTS:
		//case DISPID_PGR_DONTREDRAW:
		case DISPID_PGR_FORWARDMOUSEMESSAGESTOCONTAINEDWINDOW:
		case DISPID_PGR_HOVERTIME:
		case DISPID_PGR_RIGHTTOLEFTLAYOUT:
		case DISPID_PGR_SCROLLAUTOMATICALLY:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_PGR_BACKCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_PGR_APPID:
		case DISPID_PGR_APPNAME:
		case DISPID_PGR_APPSHORTNAME:
		case DISPID_PGR_BUILD:
		case DISPID_PGR_CHARSET:
		case DISPID_PGR_CURRENTSCROLLPOSITION:
		case DISPID_PGR_HWND:
		case DISPID_PGR_ISRELEASE:
		case DISPID_PGR_PROGRAMMER:
		case DISPID_PGR_TESTER:
		case DISPID_PGR_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_PGR_NATIVEDROPTARGET:
		case DISPID_PGR_REGISTERFOROLEDRAGDROP:
		case DISPID_PGR_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_PGR_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}

	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString Pager::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString Pager::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString Pager::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=Pager&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString Pager::GetSpecialThanks(void)
{
	return TEXT("Wine Headquarters");
}

CAtlString Pager::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString Pager::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP Pager::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_PGR_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_PGR_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_PGR_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_PGR_ORIENTATION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.orientation), description);
			break;
		case DISPID_PGR_SCROLLAUTOMATICALLY:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.scrollAutomatically), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<Pager>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP Pager::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_PGR_BORDERSTYLE:
		case DISPID_PGR_ORIENTATION:
			c = 2;
			break;
		case DISPID_PGR_APPEARANCE:
			c = 3;
			break;
		case DISPID_PGR_SCROLLAUTOMATICALLY:
			c = 4;
			break;
		case DISPID_PGR_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<Pager>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_PGR_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP Pager::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_PGR_APPEARANCE:
		case DISPID_PGR_BORDERSTYLE:
		case DISPID_PGR_MOUSEPOINTER:
		case DISPID_PGR_ORIENTATION:
		case DISPID_PGR_SCROLLAUTOMATICALLY:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<Pager>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP Pager::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	return IPerPropertyBrowsingImpl<Pager>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT Pager::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_PGR_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PGR_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PGR_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PGR_ORIENTATION:
			switch(cookie) {
				case oHorizontal:
					description = GetResStringWithNumber(IDP_ORIENTATIONHORIZONTAL, oHorizontal);
					break;
				case oVertical:
					description = GetResStringWithNumber(IDP_ORIENTATIONVERTICAL, oVertical);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_PGR_SCROLLAUTOMATICALLY:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_SCROLLAUTOMATICALLYNEVER, 0);
					break;
				case saOnMouseOver:
					description = GetResStringWithNumber(IDP_SCROLLAUTOMATICALLYONMOUSEOVER, saOnMouseOver);
					break;
				case saOnDragDrop:
					description = GetResStringWithNumber(IDP_SCROLLAUTOMATICALLYONDRAGDROP, saOnDragDrop);
					break;
				case saOnMouseOver | saOnDragDrop:
					description = GetResStringWithNumber(IDP_SCROLLAUTOMATICALLYBOTH, saOnMouseOver | saOnDragDrop);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP Pager::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 3;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StockColorPage;
		pPropertyPages->pElems[2] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP Pager::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<Pager>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP Pager::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP Pager::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT Pager::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT Pager::OnPropertyObjectChanged(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}

HRESULT Pager::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP Pager::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP Pager::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_PGR_APPEARANCE);

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_PGR_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 18;
	return S_OK;
}

STDMETHODIMP Pager::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Pager");
	return S_OK;
}

STDMETHODIMP Pager::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"PagerCtl");
	return S_OK;
}

STDMETHODIMP Pager::get_AutoScrollFrequency(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(properties.autoScrollFrequency == -1) {
			properties.autoScrollFrequency = GetDoubleClickTime() >> 3;
		}
	}

	*pValue = properties.autoScrollFrequency;
	return S_OK;
}

STDMETHODIMP Pager::put_AutoScrollFrequency(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PGR_AUTOSCROLLFREQUENCY);

	if(properties.autoScrollFrequency != newValue) {
		properties.autoScrollFrequency = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.autoScrollFrequency == -1) {
				SendMessage(PGM_SETSCROLLINFO, GetDoubleClickTime() >> 3, MAKELPARAM(properties.autoScrollLinesPerTimeout, properties.autoScrollPixelsPerLine));
			} else {
				SendMessage(PGM_SETSCROLLINFO, properties.autoScrollFrequency, MAKELPARAM(properties.autoScrollLinesPerTimeout, properties.autoScrollPixelsPerLine));
			}
		}
		FireOnChanged(DISPID_PGR_AUTOSCROLLFREQUENCY);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_AutoScrollLinesPerTimeout(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.autoScrollLinesPerTimeout;
	return S_OK;
}

STDMETHODIMP Pager::put_AutoScrollLinesPerTimeout(SHORT newValue)
{
	PUTPROPPROLOG(DISPID_PGR_AUTOSCROLLLINESPERTIMEOUT);

	if(properties.autoScrollLinesPerTimeout != newValue) {
		properties.autoScrollLinesPerTimeout = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.autoScrollFrequency == -1) {
				SendMessage(PGM_SETSCROLLINFO, GetDoubleClickTime() >> 3, MAKELPARAM(properties.autoScrollLinesPerTimeout, properties.autoScrollPixelsPerLine));
			} else {
				SendMessage(PGM_SETSCROLLINFO, properties.autoScrollFrequency, MAKELPARAM(properties.autoScrollLinesPerTimeout, properties.autoScrollPixelsPerLine));
			}
		}
		FireOnChanged(DISPID_PGR_AUTOSCROLLLINESPERTIMEOUT);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_AutoScrollPixelsPerLine(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.autoScrollPixelsPerLine;
	return S_OK;
}

STDMETHODIMP Pager::put_AutoScrollPixelsPerLine(SHORT newValue)
{
	PUTPROPPROLOG(DISPID_PGR_AUTOSCROLLPIXELSPERLINE);

	if(properties.autoScrollPixelsPerLine != newValue) {
		properties.autoScrollPixelsPerLine = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.autoScrollFrequency == -1) {
				SendMessage(PGM_SETSCROLLINFO, GetDoubleClickTime() >> 3, MAKELPARAM(properties.autoScrollLinesPerTimeout, properties.autoScrollPixelsPerLine));
			} else {
				SendMessage(PGM_SETSCROLLINFO, properties.autoScrollFrequency, MAKELPARAM(properties.autoScrollLinesPerTimeout, properties.autoScrollPixelsPerLine));
			}
		}
		FireOnChanged(DISPID_PGR_AUTOSCROLLPIXELSPERLINE);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(PGM_GETBKCOLOR, 0, 0));
		if(color == CLR_DEFAULT) {
			properties.backColor = 0x80000000 | COLOR_BTNFACE;
		} else if(color != OLECOLOR2COLORREF(properties.backColor)) {
			properties.backColor = color;
		}
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP Pager::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_PGR_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(PGM_SETBKCOLOR, 0, OLECOLOR2COLORREF(properties.backColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_PGR_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_BorderSize(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderSize = SendMessage(PGM_GETBORDER, 0, 0);
	}

	*pValue = properties.borderSize;
	return S_OK;
}

STDMETHODIMP Pager::put_BorderSize(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PGR_BORDERSIZE);

	if(properties.borderSize != newValue) {
		properties.borderSize = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(PGM_SETBORDER, 0, properties.borderSize);
			FireViewChange();
		}
		FireOnChanged(DISPID_PGR_BORDERSIZE);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP Pager::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_PGR_BORDERSTYLE);

	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_PGR_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP Pager::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP Pager::get_CurrentScrollPosition(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = SendMessage(PGM_GETPOS, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Pager::put_CurrentScrollPosition(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PGR_CURRENTSCROLLPOSITION);

	if(IsWindow()) {
		SendMessage(PGM_SETPOS, 0, newValue);
		FireViewChange();
		FireOnChanged(DISPID_PGR_CURRENTSCROLLPOSITION);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Pager::get_DetectDoubleClicks(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	return S_OK;
}

STDMETHODIMP Pager::put_DetectDoubleClicks(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PGR_DETECTDOUBLECLICKS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.detectDoubleClicks != b) {
		properties.detectDoubleClicks = b;
		SetDirty(TRUE);

		FireOnChanged(DISPID_PGR_DETECTDOUBLECLICKS);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP Pager::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_PGR_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((properties.disabledEvents & deMouseEvents) != (newValue & deMouseEvents)) {
			if(IsWindow()) {
				if(newValue & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_PGR_DISABLEDEVENTS);
	}
	return S_OK;
}

/*STDMETHODIMP Pager::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP Pager::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PGR_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_PGR_DONTREDRAW);
	}
	return S_OK;
}*/

STDMETHODIMP Pager::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP Pager::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PGR_ENABLED);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_PGR_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_ForwardMouseMessagesToContainedWindow(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.forwardMouseMessagesToContainedWindow;
	return S_OK;
}

STDMETHODIMP Pager::put_ForwardMouseMessagesToContainedWindow(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PGR_FORWARDMOUSEMESSAGESTOCONTAINEDWINDOW);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.forwardMouseMessagesToContainedWindow != b) {
		properties.forwardMouseMessagesToContainedWindow = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(PGM_FORWARDMOUSE, properties.forwardMouseMessagesToContainedWindow, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_PGR_FORWARDMOUSEMESSAGESTOCONTAINEDWINDOW);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_hContainedWindow(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(properties.hContainedWindow);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Pager::put_hContainedWindow(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_PGR_HCONTAINEDWINDOW);

	HWND hWnd = static_cast<HWND>(LongToHandle(newValue));
	if(properties.hContainedWindow != hWnd) {
		properties.hContainedWindow = hWnd;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.hContainedWindow && ::IsWindow(properties.hContainedWindow)) {
				CWindow(properties.hContainedWindow).SetParent(*this);
			}
			SendMessage(PGM_SETCHILD, 0, reinterpret_cast<LPARAM>(properties.hContainedWindow));
		}

		FireOnChanged(DISPID_PGR_HCONTAINEDWINDOW);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP Pager::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PGR_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_PGR_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP Pager::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP Pager::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP Pager::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_PGR_MOUSEICON);

	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_PGR_MOUSEICON);
	return S_OK;
}

STDMETHODIMP Pager::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_PGR_MOUSEICON);

	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_PGR_MOUSEICON);
	return S_OK;
}

STDMETHODIMP Pager::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP Pager::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_PGR_MOUSEPOINTER);

	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_PGR_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_NativeDropTarget(LPUNKNOWN* ppValue)
{
	ATLASSERT_POINTER(ppValue, LPUNKNOWN);
	if(!ppValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SendMessage(PGM_GETDROPTARGET, 0, reinterpret_cast<LPARAM>(ppValue));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Pager::get_Orientation(OrientationConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OrientationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.orientation = ((GetStyle() & PGS_HORZ) == PGS_HORZ ? oHorizontal : oVertical);
	}

	*pValue = properties.orientation;
	return S_OK;
}

STDMETHODIMP Pager::put_Orientation(OrientationConstants newValue)
{
	PUTPROPPROLOG(DISPID_PGR_ORIENTATION);
	if(properties.orientation != newValue) {
		properties.orientation = newValue;
		SetDirty(TRUE);

		RECT windowRect = m_rcPos;
		SIZEL himetric = {m_sizeExtent.cx, m_sizeExtent.cy};
		SIZEL pixels = {0};
	  AtlHiMetricToPixel(&himetric, &pixels);
		windowRect.right = windowRect.left + pixels.cy;
		windowRect.bottom = windowRect.top + pixels.cx;
		if(m_spInPlaceSite) {
			m_spInPlaceSite->OnPosRectChange(&windowRect);
		}

		if(IsWindow()) {
			switch(properties.orientation) {
				case oHorizontal:
					ModifyStyle(0, PGS_HORZ, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case oVertical:
					ModifyStyle(PGS_HORZ, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_PGR_ORIENTATION);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP Pager::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP Pager::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PGR_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				if(GetStyle() & PGS_DRAGNDROP) {
					// PGS_DRAGNDROP registers the window for drag'n'drop with the native IDropTarget
					RevokeDragDrop(*this);
				}
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			/*} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);*/
			}
		}
		FireOnChanged(DISPID_PGR_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_RightToLeftLayout(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeftLayout = ((GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.rightToLeftLayout);
	return S_OK;
}

STDMETHODIMP Pager::put_RightToLeftLayout(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PGR_RIGHTTOLEFTLAYOUT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.rightToLeftLayout != b) {
		properties.rightToLeftLayout = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.rightToLeftLayout) {
				ModifyStyleEx(0, WS_EX_LAYOUTRTL, SWP_FRAMECHANGED);
			} else {
				ModifyStyleEx(WS_EX_LAYOUTRTL, 0, SWP_FRAMECHANGED);
			}
		}
		FireOnChanged(DISPID_PGR_RIGHTTOLEFTLAYOUT);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_ScrollAutomatically(ScrollAutomaticallyConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ScrollAutomaticallyConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		DWORD style = GetStyle();
		properties.scrollAutomatically = static_cast<ScrollAutomaticallyConstants>(0);
		if(style & PGS_AUTOSCROLL) {
			properties.scrollAutomatically = static_cast<ScrollAutomaticallyConstants>(static_cast<long>(properties.scrollAutomatically) | saOnMouseOver);
		}
		if(style & PGS_DRAGNDROP) {
			properties.scrollAutomatically = static_cast<ScrollAutomaticallyConstants>(static_cast<long>(properties.scrollAutomatically) | saOnDragDrop);
		}
	}

	*pValue = properties.scrollAutomatically;
	return S_OK;
}

STDMETHODIMP Pager::put_ScrollAutomatically(ScrollAutomaticallyConstants newValue)
{
	PUTPROPPROLOG(DISPID_PGR_SCROLLAUTOMATICALLY);
	if(properties.scrollAutomatically != newValue) {
		properties.scrollAutomatically = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			DWORD newStyle = 0;
			if(properties.scrollAutomatically & saOnMouseOver) {
				newStyle |= PGS_AUTOSCROLL;
			}
			if(properties.scrollAutomatically & saOnDragDrop) {
				newStyle |= PGS_DRAGNDROP;
			}
			ModifyStyle(PGS_AUTOSCROLL | PGS_DRAGNDROP, newStyle);

			if(properties.registerForOLEDragDrop) {
				// PGS_DRAGNDROP registers the window for drag'n'drop with the native IDropTarget
				RevokeDragDrop(*this);
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			/*} else {
				// PGS_DRAGNDROP registers the window for drag'n'drop with the native IDropTarget
				RevokeDragDrop(*this);*/
			}
		}

		FireOnChanged(DISPID_PGR_SCROLLAUTOMATICALLY);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_ScrollButtonSize(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.scrollButtonSize = SendMessage(PGM_GETBUTTONSIZE, 0, 0);
	}

	*pValue = properties.scrollButtonSize;
	return S_OK;
}

STDMETHODIMP Pager::put_ScrollButtonSize(LONG newValue)
{
	PUTPROPPROLOG(DISPID_PGR_SCROLLBUTTONSIZE);

	if(properties.scrollButtonSize != newValue) {
		properties.scrollButtonSize = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(PGM_SETBUTTONSIZE, 0, properties.scrollButtonSize);
			FireViewChange();
		}
		FireOnChanged(DISPID_PGR_SCROLLBUTTONSIZE);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_ScrollButtonState(ScrollButtonConstants scrollButton, ScrollButtonStateConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ScrollButtonStateConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		*pValue = static_cast<ScrollButtonStateConstants>(SendMessage(PGM_GETBUTTONSTATE, 0, scrollButton));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Pager::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP Pager::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_PGR_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_PGR_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP Pager::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP Pager::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP Pager::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP Pager::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP Pager::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails)
{
	ATLASSERT_POINTER(pHitTestDetails, HitTestConstants);
	if(!pHitTestDetails) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pHitTestDetails = HitTest(x, y);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Pager::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP Pager::Refresh(void)
{
	if(IsWindow()) {
		//dragDropStatus.HideDragImage(TRUE);
		Invalidate();
		UpdateWindow();
		//dragDropStatus.ShowDragImage(TRUE);
	}
	return S_OK;
}

STDMETHODIMP Pager::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP Pager::UpdateScrollableSize(void)
{
	if(IsWindow()) {
		SendMessage(PGM_RECALCSIZE, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}


LRESULT Pager::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);
	return 0;
}

LRESULT Pager::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(*this);
	}
	return lr;
}

LRESULT Pager::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(*this);

	wasHandled = FALSE;
	return 0;
}

LRESULT Pager::OnEraseBkGnd(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(!properties.hContainedWindow) {
		HDC hDCDst = reinterpret_cast<HDC>(wParam);
		RECT clientRectangle = {0};
		GetClientRect(&clientRectangle);
		FillRect(hDCDst, &clientRectangle, GetSysColorBrush(COLOR_APPWORKSPACE));
	} else {
		wasHandled = FALSE;
	}
	return 1;
}

LRESULT Pager::OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Pager::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Pager::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Pager::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Pager::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Pager::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);

	return 0;
}

LRESULT Pager::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT Pager::OnNCHitTest(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	return HTCLIENT;
}

LRESULT Pager::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT Pager::OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Pager::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Pager::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	CRect clientArea;
	// NOTE: The scroll buttons are not part of the client area.
	/*GetClientRect(&clientArea);
	ClientToScreen(&clientArea);*/
	GetWindowRect(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

//LRESULT Pager::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
//{
//	if(lParam == 71216) {
//		// the message was sent by ourselves
//		lParam = 0;
//		if(wParam) {
//			// We're gonna activate redrawing - does the client allow this?
//			if(properties.dontRedraw) {
//				// no, so eat this message
//				return 0;
//			}
//		}
//	} else {
//		// TODO: Should we really do this?
//		properties.dontRedraw = !wParam;
//	}
//
//	return DefWindowProc(message, wParam, lParam);
//}

//LRESULT Pager::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
//{
//	switch(wParam) {
//		case timers.ID_REDRAW:
//			if(IsWindowVisible()) {
//				KillTimer(timers.ID_REDRAW);
//				SetRedraw(!properties.dontRedraw);
//			} else {
//				// wait... (this fixes visibility problems if another control displays a nag screen)
//			}
//			break;
//
//		default:
//			wasHandled = FALSE;
//			break;
//	}
//	return 0;
//}

LRESULT Pager::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		Invalidate();
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	if(!(pDetails->flags & SWP_NOSIZE) && properties.hContainedWindow && ::IsWindow(properties.hContainedWindow)) {
		::SetWindowPos(properties.hContainedWindow, NULL, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOREPOSITION | SWP_NOZORDER | SWP_FRAMECHANGED);
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT Pager::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT Pager::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}


LRESULT Pager::OnForwardMouse(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	properties.forwardMouseMessagesToContainedWindow = wParam;
	return lr;
}

LRESULT Pager::OnSetChild(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	properties.hContainedWindow = reinterpret_cast<HWND>(lParam);
	return lr;
}

LRESULT Pager::OnSetPos(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_PGR_CURRENTSCROLLPOSITION) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	SetDirty(TRUE);
	FireOnChanged(DISPID_PGR_CURRENTSCROLLPOSITION);
	SendOnDataChange();
	return lr;
}

LRESULT Pager::OnSetScrollInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(!IsInDesignMode()) {
		properties.autoScrollFrequency = wParam;
		properties.autoScrollLinesPerTimeout = LOWORD(lParam);
		properties.autoScrollPixelsPerLine = HIWORD(lParam);
	}
	return lr;
}


LRESULT Pager::OnCalcSizeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMPGCALCSIZE pDetails = reinterpret_cast<LPNMPGCALCSIZE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMPGCALCSIZE);
	if(!pDetails) {
		return 0;
	}

	OrientationConstants orientation = static_cast<OrientationConstants>(pDetails->dwFlag - 1);
	Raise_GetScrollableSize(orientation, reinterpret_cast<LONG*>(pDetails->dwFlag & PGF_CALCHEIGHT ? &pDetails->iHeight : &pDetails->iWidth));
	return 0;
}

LRESULT Pager::OnHotItemChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	LPNMPGHOTITEM pDetails = reinterpret_cast<LPNMPGHOTITEM>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMPGHOTITEM);
	if(!pDetails) {
		return 0;
	}

	ATLASSERT((pDetails->dwFlags & ~(HICF_ENTERING | HICF_LEAVING)) == 0 && "Unexpected flag in pDetails->dwFlags");

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!(properties.disabledEvents & deHotButtonChangeEvents)) {
		ScrollButtonConstants scrollButton = sbTopOrLeft;
		if(pDetails->dwFlags & HICF_ENTERING) {
			scrollButton = static_cast<ScrollButtonConstants>(pDetails->idNew);
		} else if(pDetails->dwFlags & HICF_LEAVING) {
			scrollButton = static_cast<ScrollButtonConstants>(pDetails->idOld);
		}

		Raise_HotButtonChanging(scrollButton, static_cast<HotButtonChangingCausedByConstants>(pDetails->dwFlags & (HICF_OTHER | HICF_MOUSE | HICF_ARROWKEYS | HICF_ACCELERATOR)), static_cast<HotButtonChangingAdditionalInfoConstants>(pDetails->dwFlags & (HICF_DUPACCEL | HICF_ENTERING | HICF_LEAVING | HICF_RESELECT | HICF_LMOUSE | HICF_TOGGLEDROPDOWN)), &cancel);
	}

	if(cancel != VARIANT_FALSE) {
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT Pager::OnScrollNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMPGSCROLL pDetails = reinterpret_cast<LPNMPGSCROLL>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMPGSCROLL);
	if(!pDetails) {
		return 0;
	}

	// NOTE: The PGK_* constants match with VB's ShiftConstants
	SHORT shift = static_cast<SHORT>(pDetails->fwKeys);
	Raise_BeforeScroll(static_cast<ScrollDirectionConstants>(pDetails->iDir), shift, reinterpret_cast<RECTANGLE*>(&pDetails->rcParent), pDetails->iXpos, pDetails->iYpos, reinterpret_cast<LONG*>(&pDetails->iScroll));
	return 0;
}


inline HRESULT Pager::Raise_BeforeScroll(ScrollDirectionConstants scrollDirection, SHORT shift, RECTANGLE* pClientRectangle, LONG horizontalScrollPosition, LONG verticalScrollPosition, LONG* pScrollDelta)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeforeScroll(scrollDirection, shift, pClientRectangle, horizontalScrollPosition, verticalScrollPosition, pScrollDelta);
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		/*if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
		}*/

		return Fire_ContextMenu(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_DestroyedControlWindow(HWND hWnd)
{
	//KillTimer(timers.ID_REDRAW);
	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_GetScrollableSize(OrientationConstants orientation, LONG* pScrollableSize)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetScrollableSize(orientation, pScrollableSize);
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_HotButtonChanging(ScrollButtonConstants scrollButton, HotButtonChangingCausedByConstants causedBy, HotButtonChangingAdditionalInfoConstants additionalInfo, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HotButtonChanging(scrollButton, causedBy, additionalInfo, pCancel);
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	BOOL fireMouseDown = FALSE;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		fireMouseDown = TRUE;
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
	}
	mouseStatus.StoreClickCandidate(button);
	// NOTE: If the child window is a tool bar, SetCapture causes problems. The pager cannot be scrolled anymore.
	//SetCapture();

	if(mouseStatus.IsDoubleClickCandidate(button)) {
		/* Watch for double-clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		// NOTE: The scroll buttons are not part of the client area.
		/*GetClientRect(&clientArea);
		ClientToScreen(&clientArea);*/
		GetWindowRect(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the double-click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_DblClick(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RDblClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MDblClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XDblClick(button, shift, x, y);
					}
					break;
			}
			mouseStatus.RemoveAllDoubleClickCandidates();
		}

		mouseStatus.RemoveClickCandidate(button);
		// NOTE: If the child window is a tool bar, SetCapture causes problems. The pager cannot be scrolled anymore.
		/*if(GetCapture() == *this) {
			ReleaseCapture();
		}*/
		fireMouseDown = FALSE;
	} else {
		mouseStatus.RemoveAllDoubleClickCandidates();
	}

	HRESULT hr = S_OK;
	if(fireMouseDown) {
		hr = Fire_MouseDown(button, shift, x, y, HitTest(x, y));
	}
	return hr;
}

inline HRESULT Pager::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseEnter(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseHover(button, shift, x, y, HitTest(x, y));
	}
	return S_OK;
}

inline HRESULT Pager::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.LeaveControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseLeave(button, shift, x, y, HitTest(x, y));
	}
	return S_OK;
}

inline HRESULT Pager::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y, HitTest(x, y));
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		// NOTE: The scroll buttons are not part of the client area.
		/*GetClientRect(&clientArea);
		ClientToScreen(&clientArea);*/
		GetWindowRect(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		// NOTE: If the child window is a tool bar, SetCapture causes problems. The pager cannot be scrolled anymore.
		/*if(GetCapture() == *this) {
			ReleaseCapture();
		}*/

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
			if(properties.detectDoubleClicks) {
				mouseStatus.StoreDoubleClickCandidate(button);
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT Pager::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition, BOOL* /*pCallDropTargetHelper*/)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = HitTest(mousePosition.x, mousePosition.y);

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
		   always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT Pager::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition, BOOL* /*pCallDropTargetHelper*/)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	HitTestConstants hitTestDetails = HitTest(mousePosition.x, mousePosition.y);

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
		}
	//}
	return hr;
}

inline HRESULT Pager::Raise_OLEDragLeave(BOOL* /*pCallDropTargetHelper*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	HitTestConstants hitTestDetails = HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, hitTestDetails);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT Pager::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition, BOOL* /*pCallDropTargetHelper*/)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HitTestConstants hitTestDetails = HitTest(mousePosition.x, mousePosition.y);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
		}
	//}
	return hr;
}

inline HRESULT Pager::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_RecreatedControlWindow(HWND hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		if(GetStyle() & PGS_DRAGNDROP) {
			// PGS_DRAGNDROP registers the window for drag'n'drop with the native IDropTarget
			RevokeDragDrop(*this);
		}
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	/*} else {
		// PGS_DRAGNDROP registers the window for drag'n'drop with the native IDropTarget
		RevokeDragDrop(*this);*/
	}

	/*if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}*/

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}

inline HRESULT Pager::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y, HitTest(x, y));
	//}
	//return S_OK;
}


void Pager::RecreateControlWindow(void)
{
	// This method shouldn't be used, because it will destroy all contained controls.
	ATLASSERT(FALSE);
	/*if(m_bInPlaceActive && !flags.dontRecreate) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}*/
}

DWORD Pager::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeftLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	return extendedStyle;
}

DWORD Pager::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(properties.scrollAutomatically & saOnMouseOver) {
		style |= PGS_AUTOSCROLL;
	}
	if(properties.scrollAutomatically & saOnDragDrop) {
		style |= PGS_DRAGNDROP;
	}
	switch(properties.orientation) {
		case oHorizontal:
			style |= PGS_HORZ;
			break;
	}
	return style;
}

void Pager::SendConfigurationMessages(void)
{
	SendMessage(PGM_SETBKCOLOR, 0, OLECOLOR2COLORREF(properties.backColor));
	SendMessage(PGM_SETBUTTONSIZE, 0, properties.scrollButtonSize);
	// NOTE: Border size is limited by button size.
	SendMessage(PGM_SETBORDER, 0, properties.borderSize);
	if(properties.autoScrollFrequency >= 0) {
		SendMessage(PGM_SETSCROLLINFO, properties.autoScrollFrequency, MAKELPARAM(properties.autoScrollLinesPerTimeout, properties.autoScrollPixelsPerLine));
	} else if(properties.autoScrollLinesPerTimeout > 0) {
		SendMessage(PGM_SETSCROLLINFO, GetDoubleClickTime() >> 3, MAKELPARAM(properties.autoScrollLinesPerTimeout, properties.autoScrollPixelsPerLine));
	}
	SendMessage(PGM_FORWARDMOUSE, properties.forwardMouseMessagesToContainedWindow, 0);
}

HCURSOR Pager::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

HitTestConstants Pager::HitTest(LONG x, LONG y)
{
	ATLASSERT(IsWindow());

	POINT pt = {x, y};
	HitTestConstants flags = static_cast<HitTestConstants>(0);

	CRect windowRectangle;
	GetWindowRect(&windowRectangle);
	ScreenToClient(&windowRectangle);
	if(windowRectangle.PtInRect(pt)) {
		// Are we inside the client area?
		CRect clientRectangle;
		GetClientRect(&clientRectangle);
		if(clientRectangle.PtInRect(pt)) {
			// bingo
			BOOL isVertical = !(GetStyle() & PGS_HORZ);
			DWORD borderSize = SendMessage(PGM_GETBORDER, 0, 0);
			if(isVertical) {
				if(x < clientRectangle.left + static_cast<LONG>(borderSize)) {
					flags = static_cast<HitTestConstants>(flags | htTopOrLeftBorder);
				} else if(x >= clientRectangle.right - static_cast<LONG>(borderSize)) {
					flags = static_cast<HitTestConstants>(flags | htBottomOrRightBorder);
				} else {
					flags = static_cast<HitTestConstants>(flags | htClientArea);
				}
			} else {
				if(y < clientRectangle.top + static_cast<LONG>(borderSize)) {
					flags = static_cast<HitTestConstants>(flags | htTopOrLeftBorder);
				} else if(y >= clientRectangle.bottom - static_cast<LONG>(borderSize)) {
					flags = static_cast<HitTestConstants>(flags | htBottomOrRightBorder);
				} else {
					flags = static_cast<HitTestConstants>(flags | htClientArea);
				}
			}
		} else {
			// we are either over a border or over a scroll button
			BOOL isVertical = !(GetStyle() & PGS_HORZ);
			DWORD buttonSize = SendMessage(PGM_GETBUTTONSIZE, 0, 0);
			if(isVertical) {
				if(y < windowRectangle.top + static_cast<LONG>(buttonSize)) {
					flags = static_cast<HitTestConstants>(flags | htTopOrLeftScrollButton);
				} else if(y >= windowRectangle.bottom - static_cast<LONG>(buttonSize)) {
					flags = static_cast<HitTestConstants>(flags | htBottomOrRightScrollButton);
				}
			} else {
				if(x < windowRectangle.left + static_cast<LONG>(buttonSize)) {
					flags = static_cast<HitTestConstants>(flags | htTopOrLeftScrollButton);
				} else if(x >= windowRectangle.right - static_cast<LONG>(buttonSize)) {
					flags = static_cast<HitTestConstants>(flags | htBottomOrRightScrollButton);
				}
			}
		}
	} else {
		if(x < windowRectangle.left) {
			flags = static_cast<HitTestConstants>(flags | htToLeft);
		} else if(x >= windowRectangle.right) {
			flags = static_cast<HitTestConstants>(flags | htToRight);
		}
		if(y < windowRectangle.top) {
			flags = static_cast<HitTestConstants>(flags | htAbove);
		} else if(y >= windowRectangle.bottom) {
			flags = static_cast<HitTestConstants>(flags | htBelow);
		}
	}
	return flags;
}

BOOL Pager::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}