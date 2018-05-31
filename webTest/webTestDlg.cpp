
// webTestDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "webTest.h"
#include "webTestDlg.h"
#include "afxdialogex.h"
#import <mshtml.tlb>  
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CwebTestDlg 对话框

BEGIN_DHTML_EVENT_MAP(CwebTestDlg)
	DHTML_EVENT_ONCLICK(_T("ButtonOK"), OnButtonOK)
	DHTML_EVENT_ONCLICK(_T("ButtonCancel"), OnButtonCancel)
	//ON_EVENT(CwebTestDlg, AFX_IDC_BROWSER, 250 /* BeforeNavigate2 */, _OnBeforeNavigate2, VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PBOOL)
END_DHTML_EVENT_MAP()



CwebTestDlg::CwebTestDlg(CWnd* pParent /*=NULL*/)
	: CDHtmlDialog(CwebTestDlg::IDD, CwebTestDlg::IDH, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	//  m_webbrowser = 0;
	showRes=0;
}

void CwebTestDlg::DoDataExchange(CDataExchange* pDX)
{
	CDHtmlDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EXPLORER1, m_webbrowser);
}

BEGIN_MESSAGE_MAP(CwebTestDlg, CDHtmlDialog)
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(IDC_BUTTON1, &CwebTestDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CwebTestDlg 消息处理程序

BOOL CwebTestDlg::OnInitDialog()
{
	CDHtmlDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	//LoadFromResource(IDR_HTML1);//ID_XXX 是资源的定义
	char self_path[MAX_PATH] = { 0 };
	//CString self_path="";
	GetModuleFileName(NULL, self_path, MAX_PATH);
	CString res_url;
	//res_url.Format("res://%s", "C:\\Users\\admin\\Desktop\\webTest\\webTest\\HTMLPage1.htm");
	res_url.Format("res://%s/%d", self_path, IDR_HTML1);
	m_webbrowser.Navigate(res_url, NULL, NULL, NULL, NULL);
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CwebTestDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDHtmlDialog::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CwebTestDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDHtmlDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CwebTestDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

HRESULT CwebTestDlg::OnButtonOK(IHTMLElement* /*pElement*/)
{
	OnOK();
	return S_OK;
}

HRESULT CwebTestDlg::OnButtonCancel(IHTMLElement* /*pElement*/)
{
	OnCancel();
	return S_OK;
}


BOOL CwebTestDlg::PreTranslateMessage(MSG* pMsg)
{
	// TODO: 在此添加专用代码和/或调用基类
	if ((pMsg->message == WM_RBUTTONDOWN) ||
		(pMsg->message == WM_RBUTTONDBLCLK))
		return TRUE;
	else
		//return CHtmlView::PreTranslateMessage(pMsg);
		return CDHtmlDialog::PreTranslateMessage(pMsg);
}


//void CwebTestDlg::OnBeforeNavigate(LPDISPATCH pDisp, LPCTSTR szUrl)
//{
//	// TODO: 在此添加专用代码和/或调用基类
//	if (pDisp == FrameC && lstrcmpW(V_BSTR(url), L"about:blank")!=0 ) 
//	{ 
//		*Cancel = VARIANT_TRUE; 
//		CComQIPtr <IWebBrowser2> spBrowser =  pDisp; 
//		CComVariant vUrl=L"about:blank", vTmp; 
//		spBrowser->Navigate2(&vUrl, &vTmp, &vTmp, &vTmp, &vTmp); 
//	} 
//	CDHtmlDialog::OnBeforeNavigate(pDisp, szUrl);
//}

//void CDHtmlDialog::_OnBeforeNavigate2(LPDISPATCH pDisp, VARIANT FAR* URL, VARIANT FAR* Flags, VARIANT FAR* TargetFrameName, VARIANT FAR* PostData, VARIANT FAR* Headers, BOOL FAR* Cancel)
//{
//	Flags; // unused
//	TargetFrameName; // unused
//	PostData; // unused
//	Headers; // unused
//	Cancel; // unused
//
//	CString str(V_BSTR(URL));
//	OnBeforeNavigate(pDisp, str);
//}
//void CwebTestDlg::_OnBeforeNavigate2(LPDISPATCH pDisp, VARIANT FAR* URL, VARIANT FAR* Flags, VARIANT FAR* TargetFrameName, VARIANT FAR* PostData, VARIANT FAR* Headers, BOOL FAR* Cancel)
//{
//	Flags; // unused
//	TargetFrameName; // unused
//	PostData; // unused
//	Headers; // unused
//	//Cancel; // unused
//	*Cancel = TRUE; //这里非零就可以阻止进一步操作
//
//	CString str(V_BSTR(URL));
//	OnBeforeNavigate(pDisp, str);
//}
//0x00b69f10 
void CwebTestDlg::OnBeforeNavigate(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	// TODO: 在此添加专用代码和/或调用基类
	CString tmp="";
	tmp.Format("%s",szUrl);
	int index=tmp.Find( '#' );
	if (index == -1 )
	{

	}
	else
	{
		//int tmpI=atoi(tmp.Left(index-1));
		//if (tmpI == IDR_HTML1)
		//{
			CString aCommand=tmp.Mid(index+1);
			if (!aCommand.Compare("A"))
			{
				OnBnClickedButton1();
			}
		//}
	}
	
	CDHtmlDialog::OnBeforeNavigate(pDisp, szUrl);
}

bool CwebTestDlg::ProcessHtmlCommand(LPCTSTR lpszURL)
{
	CString str = lpszURL;
	AfxMessageBox(str);
	if(-1 != str.Find("fileDlg.open")){
		//::AfxMessageBox("fileDlg.open");
		CFileDialog fd(true);
		if(IDOK == fd.DoModal())AfxMessageBox(fd.GetFileName());
		return true;
	}
	return false;
}

void CwebTestDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	CComQIPtr<IHTMLDocument2> document = m_webbrowser.get_Document();  
	CComDispatchDriver script;  
	document->get_Script(&script);  
	CComVariant var(static_cast<IDispatch*>(this));  
	script.Invoke1(L"SaveCppObject", &var);  
	//CString tmpStr="";
	//showRes++;
	//tmpStr.Format("data:%d",showRes);
	//GetDlgItem(IDC_BUTTON1)->SetWindowText(tmpStr);
}

DWORD CwebTestDlg::GetProcessID()  
{  
	return GetCurrentProcessId();  
}  

DWORD CwebTestDlg::test1()  
{  
	CString tmpStr="";
	showRes++;
	tmpStr.Format("data:%d",showRes);
	GetDlgItem(IDC_BUTTON1)->SetWindowText(tmpStr);
	return showRes;
}  

void CwebTestDlg::ShowMessageBox(const wchar_t *msg)  
{  
	CString tmpStr(msg);
	//msg=tmpStr.AllocSysString();
	MessageBox(tmpStr,"这是来自javascript的消息");  

}  

//我自己给我的两个函数拟定的数字ID，这个ID可以取0-16384之间的任意数
enum
{
	FUNCTION_ShowMessageBox = 1,
	FUNCTION_GetProcessID = 2,
	FUNCTION_test1=3,
};

//不用实现，直接返回E_NOTIMPL
HRESULT STDMETHODCALLTYPE CwebTestDlg::GetTypeInfoCount(UINT *pctinfo)
{
	return E_NOTIMPL;
}

//不用实现，直接返回E_NOTIMPL
HRESULT STDMETHODCALLTYPE CwebTestDlg::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

//JavaScript调用这个对象的方法时，会把方法名，放到rgszNames中，我们需要给这个方法名拟定一个唯一的数字ID，用rgDispId传回给它
//同理JavaScript存取这个对象的属性时，会把属性名放到rgszNames中，我们需要给这个属性名拟定一个唯一的数字ID，用rgDispId传回给它
//紧接着JavaScript会调用Invoke，并把这个ID作为参数传递进来
HRESULT STDMETHODCALLTYPE CwebTestDlg::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	//rgszNames是个字符串数组，cNames指明这个数组中有几个字符串，如果不是1个字符串，忽略它
	if (cNames != 1)
		return E_NOTIMPL;
	//如果字符串是ShowMessageBox，说明JavaScript在调用我这个对象的ShowMessageBox方法，我就把我拟定的ID通过rgDispId告诉它
	if (wcscmp(rgszNames[0], L"ShowMessageBox") == 0)
	{
		*rgDispId = FUNCTION_ShowMessageBox;
		return S_OK;
	}
	//同理，如果字符串是GetProcessID，说明JavaScript在调用我这个对象的GetProcessID方法
	else if (wcscmp(rgszNames[0], L"GetProcessID") == 0)
	{
		*rgDispId = FUNCTION_GetProcessID;
		return S_OK;
	}
	else if (wcscmp(rgszNames[0], L"test1") == 0)
	{
		*rgDispId = FUNCTION_test1;
		return S_OK;
	}
	
	else
		return E_NOTIMPL;
}

//JavaScript通过GetIDsOfNames拿到我的对象的方法的ID后，会调用Invoke，dispIdMember就是刚才我告诉它的我自己拟定的ID
//wFlags指明JavaScript对我的对象干了什么事情！
//如果是DISPATCH_METHOD，说明JavaScript在调用这个对象的方法，比如cpp_object.ShowMessageBox();
//如果是DISPATCH_PROPERTYGET，说明JavaScript在获取这个对象的属性，比如var n = cpp_object.num;
//如果是DISPATCH_PROPERTYPUT，说明JavaScript在修改这个对象的属性，比如cpp_object.num = 10;
//如果是DISPATCH_PROPERTYPUTREF，说明JavaScript在通过引用修改这个对象，具体我也不懂
//示例代码并没有涉及到wFlags和对象属性的使用，需要的请自行研究，用法是一样的
//pDispParams就是JavaScript调用我的对象的方法时传递进来的参数，里面有一个数组保存着所有参数
//pDispParams->cArgs就是数组中有多少个参数
//pDispParams->rgvarg就是保存着参数的数组，请使用[]下标来访问，每个参数都是VARIANT类型，可以保存各种类型的值
//具体是什么类型用VARIANT::vt来判断，不多解释了，VARIANT这东西大家都懂
//pVarResult就是我们给JavaScript的返回值
//其它不用管
HRESULT STDMETHODCALLTYPE CwebTestDlg::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
	WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	//通过ID我就知道JavaScript想调用哪个方法
	if (dispIdMember == FUNCTION_ShowMessageBox)
	{
		//检查是否只有一个参数
		if (pDispParams->cArgs != 1)
			return E_NOTIMPL;
		//检查这个参数是否是字符串类型
		if (pDispParams->rgvarg[0].vt != VT_BSTR)
			return E_NOTIMPL;
		//放心调用
		ShowMessageBox(pDispParams->rgvarg[0].bstrVal);
		return S_OK;
	}
	else if (dispIdMember == FUNCTION_GetProcessID)
	{
		DWORD id = GetProcessID();
		*pVarResult = CComVariant(id);
		return S_OK;
	}
	else if (dispIdMember == FUNCTION_test1)
	{
		DWORD id = test1();
		*pVarResult = CComVariant(id);
		return S_OK;
	}
	else
		return E_NOTIMPL;
}

//JavaScript拿到我们传递给它的指针后，由于它不清楚我们的对象是什么东西，会调用QueryInterface来询问我们“你是什么鬼东西？”
//它会通过riid来问我们是什么东西，只有它问到我们是不是IID_IDispatch或我们是不是IID_IUnknown时，我们才能肯定的回答它S_OK
//因为我们的对象继承于IDispatch，而IDispatch又继承于IUnknown，我们只实现了这两个接口，所以只能这样来回答它的询问
HRESULT STDMETHODCALLTYPE CwebTestDlg::QueryInterface(REFIID riid, void **ppvObject)
{
	if (riid == IID_IDispatch || riid == IID_IUnknown)
	{
		//对的，我是一个IDispatch，把我自己(this)交给你
		*ppvObject = static_cast<IDispatch*>(this);
		return S_OK;
	}
	else
		return E_NOINTERFACE;
}

//我们知道COM对象使用引用计数来管理对象生命周期，我们的CJsCallCppDlg对象的生命周期就是整个程序的生命周期
//我的这个对象不需要你JavaScript来管，我自己会管，所以我不用实现AddRef()和Release()，这里乱写一些。
//你要return 1;return 2;return 3;return 4;return 5;都可以
ULONG STDMETHODCALLTYPE CwebTestDlg::AddRef()
{
	return 1;
}

//同上，不多说了
//题外话：当然如果你要new出一个c++对象来并扔给JavaScript来管，你就需要实现AddRef()和Release()，在引用计数归零时delete this;
ULONG STDMETHODCALLTYPE CwebTestDlg::Release()
{
	return 1;
}BEGIN_EVENTSINK_MAP(CwebTestDlg, CDHtmlDialog)
	ON_EVENT(CwebTestDlg, IDC_EXPLORER1, 250, CwebTestDlg::BeforeNavigate2Explorer1, VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PBOOL)
	ON_EVENT(CwebTestDlg, IDC_EXPLORER1, 259, CwebTestDlg::DocumentCompleteExplorer1, VTS_DISPATCH VTS_PVARIANT)
	END_EVENTSINK_MAP()


void CwebTestDlg::BeforeNavigate2Explorer1(LPDISPATCH pDisp, VARIANT* URL, VARIANT* Flags, VARIANT* TargetFrameName, VARIANT* PostData, VARIANT* Headers, BOOL* Cancel)
{
	// TODO: 在此处添加消息处理程序代码
}


//该事件响应函数会在WebBrowser加载HTML文档完毕后被调用
//此时调用JavaScript的SaveCppObject函数，把我自己(this)交给它，SaveCppObject会把我这个对象保存到全局变量var cpp_object;中
//以后JavaScript就可以通过cpp_object来调用我这个C++对象的方法了
//注意，因为我的程序只用m_webbrowser.Navigate加载了一个HTML，所以这里我没有检查VARIANT* URL
//如果你的程序中有多次m_webbrowser.Navigate，需要通过VARIANT* URL判断这次事件响应是对应着哪次Navigate
void CwebTestDlg::DocumentCompleteExplorer1(LPDISPATCH pDisp, VARIANT* URL)
{
	CComQIPtr<IHTMLDocument2> document = m_webbrowser.get_Document();
	CComDispatchDriver script;
	document->get_Script(&script);
	CComVariant var(static_cast<IDispatch*>(this));
	script.Invoke1(L"SaveCppObject", &var);
}
