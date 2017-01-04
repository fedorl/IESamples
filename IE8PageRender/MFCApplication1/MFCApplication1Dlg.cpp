
// MFCApplication1Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "MFCApplication1.h"
#include "MFCApplication1Dlg.h"
#include "afxdialogex.h"
#include <Mshtml.h>
#include <sstream>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CMFCApplication1Dlg dialog



CMFCApplication1Dlg::CMFCApplication1Dlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CMFCApplication1Dlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMFCApplication1Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EXPLORER1, m_browser);
}

BEGIN_MESSAGE_MAP(CMFCApplication1Dlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CMFCApplication1Dlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CMFCApplication1Dlg::OnBnClickedButton2)
END_MESSAGE_MAP()


// CMFCApplication1Dlg message handlers

BOOL CMFCApplication1Dlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	m_browser.put_Silent(TRUE);

	OnBnClickedButton2();

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CMFCApplication1Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMFCApplication1Dlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMFCApplication1Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void convertEmfToBitmap(const RECT& fitRect, HDC hTargetDC, HENHMETAFILE hMetafile, LPCTSTR fileName);
CComPtr<IHTMLDOMNode> appendPadElement(IHTMLDocument2* pDoc, IHTMLElement* pBody, long left, long top, long width, long height);
void removeElement(IHTMLElement* pParent, IHTMLDOMNode* pChild);

void CMFCApplication1Dlg::OnBnClickedButton2()
{
	COleVariant varNull;
	COleVariant varUrl = L"http://www.bing.com/search?q=ie+must+die&qs=n";
	m_browser.Navigate2(varUrl, varNull, varNull, varNull, varNull);
}


void CMFCApplication1Dlg::OnBnClickedButton1()
{
	//get html interfaces
	IDispatch* pHtmlDoc = m_browser.get_Document();
	CComPtr<IHTMLDocument2> pHtmlDocument2;
	pHtmlDoc->QueryInterface(IID_IHTMLDocument2, (void**)&pHtmlDocument2);

	CComPtr<IHTMLElement> pBody;
	pHtmlDocument2->get_body(&pBody);

	CComPtr<IHTMLElement2> pBody2;
	pBody->QueryInterface(IID_IHTMLElement2, (void**)&pBody2);

	CComPtr<IHTMLBodyElement> pBodyElement;
	pBody->QueryInterface(IID_IHTMLBodyElement, (void**)&pBodyElement);

	CComPtr<IHTMLElement> pHtml;
	pBody->get_parentElement(&pHtml);

	CComPtr<IHTMLElement2> pHtml2;
	pHtml->QueryInterface(IID_IHTMLElement2, (void**)&pHtml2);

	CComPtr<IHTMLStyle> pHtmlStyle;
	pHtml->get_style(&pHtmlStyle);
	CComPtr<IHTMLStyle> pBodyStyle;
	pBody->get_style(&pBodyStyle);

	//get screen info
	HDC hWndDc = ::GetDC(m_hWnd);
	const int wndLogPx = GetDeviceCaps(hWndDc, LOGPIXELSX);
	const int wndLogPy = GetDeviceCaps(hWndDc, LOGPIXELSY);


	//keep current values
	SIZE keptBrowserSize = { m_browser.get_Width(), m_browser.get_Height() };
	SIZE keptScrollPos;
	//set reasonable viewport size 
	//m_browser.put_Width(docSize.cx);
	//m_browser.put_Height(docSize.cy*2);
	pHtml2->get_scrollLeft(&keptScrollPos.cx);
	pHtml2->get_scrollTop(&keptScrollPos.cy);
	COleVariant keptOverflow;
	pBodyStyle->get_overflow(&keptOverflow.bstrVal);

	//setup style and hide scroll bars
	pHtmlStyle->put_border(L"0px;");
	pHtmlStyle->put_overflow(L"hidden");
	//pBodyStyle->put_border(L"0px;");
	//pBodyStyle->put_overflow(L"hidden");

	//get document size and visible area in screen pixels
	SIZE docSize;
	pBody2->get_scrollWidth(&docSize.cx);
	pBody2->get_scrollHeight(&docSize.cy);
	RECT clientRect = { 0 };
	pHtml2->get_clientWidth(&clientRect.right);
	pHtml2->get_clientHeight(&clientRect.bottom);

	//derive chunk size
	const SIZE clientChunkSize = { 
		clientRect.right - clientRect.right % wndLogPx, 
		clientRect.bottom - clientRect.bottom % wndLogPy };

	//pad with absolutely positioned element to have enough scroll area for all chunks
	//alternatively, browser can be resized to chunk multiplies (simplest), to DPI multiplies (more work). 
	//This pad also can be made smaller, to modulus DPI, but then need more work in the loops below
	CComPtr<IHTMLDOMNode> pPadNode = 
		appendPadElement(pHtmlDocument2, pBody, docSize.cx, docSize.cy, clientChunkSize.cx, clientChunkSize.cy);

	//get printer info
	CPrintDialog pd(TRUE, PD_ALLPAGES | PD_USEDEVMODECOPIES | PD_NOPAGENUMS | PD_HIDEPRINTTOFILE | PD_NOSELECTION);
	pd.m_pd.Flags |= PD_RETURNDC | PD_RETURNDEFAULT;
	pd.DoModal(); 
	HDC hPrintDC = pd.CreatePrinterDC();
	const int printLogPx = GetDeviceCaps(hPrintDC, LOGPIXELSX);
	const int printLogPy = GetDeviceCaps(hPrintDC, LOGPIXELSY);
	const int printHorRes = ::GetDeviceCaps(hPrintDC, HORZRES);
	const SIZE printChunkSize = { printLogPx * clientChunkSize.cx / wndLogPx, printLogPy * clientChunkSize.cy / wndLogPy };

	//browser total unscaled print area in printer pixel space
	const RECT printRectPx = { 0, 0, docSize.cx* printLogPx / wndLogPx, docSize.cy*printLogPy / wndLogPy };
	//unscaled target EMF size in 0.01 mm with printer resolution
	const RECT outRect001Mm = { 0, 0, 2540 * docSize.cx / wndLogPx, 2540 * docSize.cy / wndLogPy };
	HDC hMetaDC = CreateEnhMetaFile(hPrintDC, L"c:\\temp\\test.emf", &outRect001Mm, NULL);
	::FillRect(hMetaDC, &printRectPx, (HBRUSH)::GetStockObject(BLACK_BRUSH));

	//unscaled chunk EMF size in pixels with printer resolution
	const RECT chunkRectPx = { 0, 0, printChunkSize.cx, printChunkSize.cy };
	//unscaled chunk EMF size in 0.01 mm with printer resolution
	const RECT chunkRect001Mm = { 0, 0, 2540 * clientChunkSize.cx / wndLogPx, 2540 * clientChunkSize.cy / wndLogPy };

	////////
	//render page content to metafile by small chunks

	//get renderer interface
	CComPtr<IHTMLElementRender> pRender;
	pHtml->QueryInterface(IID_IHTMLElementRender, (void**)&pRender);
	COleVariant printName = L"EMF";
	pRender->SetDocumentPrinter(printName.bstrVal, hMetaDC);


	//current positions and target area
	RECT chunkDestRectPx = { 0, 0, printChunkSize.cx, printChunkSize.cy };
	POINT clientPos = { 0, 0 };
	POINT printPos = { 0, 0 };

	//loop over chunks left to right top to bottom until scroll area is completely covered
	const SIZE lastScroll = { docSize.cx, docSize.cy};
	while (clientPos.y < lastScroll.cy)
	{
		while (clientPos.x < lastScroll.cx)
		{
			//update horizontal scroll position and set target area
			pHtml2->put_scrollLeft(clientPos.x);
			chunkDestRectPx.left = printPos.x;
			chunkDestRectPx.right = printPos.x + printChunkSize.cx;

			//render to new emf, can be optimized to avoid recreation
			HDC hChunkDC = CreateEnhMetaFile(hPrintDC, NULL, &chunkRect001Mm, NULL);
			::FillRect(hChunkDC, &chunkRectPx, (HBRUSH)::GetStockObject(WHITE_BRUSH));
			pRender->DrawToDC(hChunkDC);
			HENHMETAFILE hChunkMetafile = CloseEnhMetaFile(hChunkDC);

			//copy chunk to the main metafile
			PlayEnhMetaFile(hMetaDC, hChunkMetafile, &chunkDestRectPx);
			DeleteEnhMetaFile(hChunkMetafile);

			//update horizontal positions
			clientPos.x += clientChunkSize.cx;
			printPos.x += printChunkSize.cx;
		}

		//reset horizontal positions
		clientPos.x = 0;
		printPos.x = 0;
		//update vertical positions
		clientPos.y += clientChunkSize.cy;
		printPos.y += printChunkSize.cy;
		pHtml2->put_scrollTop(clientPos.y);
		chunkDestRectPx.top = printPos.y;
		chunkDestRectPx.bottom = printPos.y + printChunkSize.cy;
	}

	//restore changed values on browser
	//if for large pages on slow PC you get content scrolling during rendering and it is a problem,
	//you can either hide the browser and show "working" or place on top first chunk content
	pHtmlStyle->put_overflow(keptOverflow.bstrVal);
	pHtml2->put_scrollLeft(keptScrollPos.cx);
	pHtml2->put_scrollTop(keptScrollPos.cy);
	m_browser.put_Width(keptBrowserSize.cx);
	m_browser.put_Height(keptBrowserSize.cy);
	removeElement(pBody, pPadNode);

	//draw to bitmap and close metafile
	HENHMETAFILE hMetafile = CloseEnhMetaFile(hMetaDC);
	RECT fitRect = { 0, 0, printHorRes, docSize.cy * printHorRes / docSize.cx };
	convertEmfToBitmap(fitRect, hWndDc, hMetafile, L"c:\\temp\\test.bmp");
	DeleteEnhMetaFile(hMetafile);

	//cleanup - probably more here
	::ReleaseDC(m_hWnd, hWndDc);
	::DeleteDC(hPrintDC);

	//{
	//	std::stringstream ss;
	//	ss << "====" << docSize.cx << "x" << docSize.cy << " -> " << fitRect.right << "x" << fitRect.bottom << "" << "\n";
	//	OutputDebugStringA(ss.str().c_str());
	//}

}



///////////////
////some util 

void convertEmfToBitmap(const RECT& fitRect, HDC hTargetDC, HENHMETAFILE hMetafile, LPCTSTR fileName)
{
	//Create memory DC to render into
	HDC hCompDc = ::CreateCompatibleDC(hTargetDC);
	//NOTE this 
	BITMAPINFOHEADER infoHeader = { 0 };
	infoHeader.biSize = sizeof(infoHeader);
	infoHeader.biWidth = fitRect.right;
	infoHeader.biHeight = -fitRect.bottom;
	infoHeader.biPlanes = 1;
	infoHeader.biBitCount = 24;
	infoHeader.biCompression = BI_RGB;

	BITMAPINFO info;
	info.bmiHeader = infoHeader;

	//create bitmap
	BYTE* pMemory = 0;
	HBITMAP hBitmap = ::CreateDIBSection(hCompDc, &info, DIB_RGB_COLORS, (void**)&pMemory, 0, 0);
	HBITMAP hOldBmp = (HBITMAP)::SelectObject(hCompDc, hBitmap);


	PlayEnhMetaFile(hCompDc, hMetafile, &fitRect);

	BITMAPFILEHEADER fileHeader = { 0 };
	fileHeader.bfType = 0x4d42;
	fileHeader.bfSize = 0;
	fileHeader.bfReserved1 = 0;
	fileHeader.bfReserved2 = 0;
	fileHeader.bfOffBits = sizeof(BITMAPFILEHEADER)+sizeof(BITMAPINFOHEADER);

	CFile file(
		fileName,
		CFile::modeCreate | CFile::modeReadWrite | CFile::shareDenyNone);
	file.Write((char*)&fileHeader, sizeof(fileHeader));
	file.Write((char*)&infoHeader, sizeof(infoHeader));

	int bytes = (((24 * infoHeader.biWidth + 31) & (~31)) / 8) * abs(infoHeader.biHeight);
	file.Write(pMemory, bytes);

	::SelectObject(hCompDc, hOldBmp);

	//Clean up
	if (hBitmap)
		::DeleteObject(hBitmap);
	::DeleteDC(hCompDc);
}


CComPtr<IHTMLDOMNode> appendPadElement(IHTMLDocument2* pDoc, IHTMLElement* pBody, long left, long top, long width, long height)
{
	CComPtr<IHTMLElement> pPadElement;
	pDoc->createElement(L"DIV", &pPadElement);
	CComPtr<IHTMLStyle> pPadStyle;
	pPadElement->get_style(&pPadStyle);
	CString padHtml;
	const int padLineCount = max(height / 5, 1);
	for (int i = 0; i < padLineCount; i++)
	{
		padHtml.Append(L"<br> &nbsp;");
	}
	pPadElement->put_innerHTML(padHtml.GetBuffer());
	CComVariant value = left+width+1;
	pPadStyle->put_paddingLeft(value);
	pPadStyle->put_paddingRight(value);
	//pPadStyle->put_backgroundColor(CComVariant("red"));

	CComPtr<IHTMLDOMNode> pPadNode;
	pPadElement->QueryInterface(IID_IHTMLDOMNode, (void**)&pPadNode);
	CComPtr<IHTMLDOMNode> pBodyNode;
	pBody->QueryInterface(IID_IHTMLDOMNode, (void **)&pBodyNode);
	pBodyNode->appendChild(pPadNode, NULL);
	return pPadNode;
}

void removeElement(IHTMLElement* pParent, IHTMLDOMNode* pChild)
{
	CComPtr<IHTMLDOMNode> pNode;
	pParent->QueryInterface(IID_IHTMLDOMNode, (void **)&pNode);
	pNode->removeChild(pChild, NULL);
}


