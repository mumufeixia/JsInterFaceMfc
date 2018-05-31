
// webTestDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "webTest.h"
#include "webTestDlg.h"
#include "afxdialogex.h"
#import <mshtml.tlb>  
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CwebTestDlg �Ի���

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


// CwebTestDlg ��Ϣ�������

BOOL CwebTestDlg::OnInitDialog()
{
	CDHtmlDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	//LoadFromResource(IDR_HTML1);//ID_XXX ����Դ�Ķ���
	char self_path[MAX_PATH] = { 0 };
	//CString self_path="";
	GetModuleFileName(NULL, self_path, MAX_PATH);
	CString res_url;
	//res_url.Format("res://%s", "C:\\Users\\admin\\Desktop\\webTest\\webTest\\HTMLPage1.htm");
	res_url.Format("res://%s/%d", self_path, IDR_HTML1);
	m_webbrowser.Navigate(res_url, NULL, NULL, NULL, NULL);
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CwebTestDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDHtmlDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
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
	// TODO: �ڴ����ר�ô����/����û���
	if ((pMsg->message == WM_RBUTTONDOWN) ||
		(pMsg->message == WM_RBUTTONDBLCLK))
		return TRUE;
	else
		//return CHtmlView::PreTranslateMessage(pMsg);
		return CDHtmlDialog::PreTranslateMessage(pMsg);
}


//void CwebTestDlg::OnBeforeNavigate(LPDISPATCH pDisp, LPCTSTR szUrl)
//{
//	// TODO: �ڴ����ר�ô����/����û���
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
//	*Cancel = TRUE; //�������Ϳ�����ֹ��һ������
//
//	CString str(V_BSTR(URL));
//	OnBeforeNavigate(pDisp, str);
//}
//0x00b69f10 
void CwebTestDlg::OnBeforeNavigate(LPDISPATCH pDisp, LPCTSTR szUrl)
{
	// TODO: �ڴ����ר�ô����/����û���
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
	// TODO: �ڴ���ӿؼ�֪ͨ����������
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
	MessageBox(tmpStr,"��������javascript����Ϣ");  

}  

//���Լ����ҵ����������ⶨ������ID�����ID����ȡ0-16384֮���������
enum
{
	FUNCTION_ShowMessageBox = 1,
	FUNCTION_GetProcessID = 2,
	FUNCTION_test1=3,
};

//����ʵ�֣�ֱ�ӷ���E_NOTIMPL
HRESULT STDMETHODCALLTYPE CwebTestDlg::GetTypeInfoCount(UINT *pctinfo)
{
	return E_NOTIMPL;
}

//����ʵ�֣�ֱ�ӷ���E_NOTIMPL
HRESULT STDMETHODCALLTYPE CwebTestDlg::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

//JavaScript�����������ķ���ʱ����ѷ��������ŵ�rgszNames�У�������Ҫ������������ⶨһ��Ψһ������ID����rgDispId���ظ���
//ͬ��JavaScript��ȡ������������ʱ������������ŵ�rgszNames�У�������Ҫ������������ⶨһ��Ψһ������ID����rgDispId���ظ���
//������JavaScript�����Invoke���������ID��Ϊ�������ݽ���
HRESULT STDMETHODCALLTYPE CwebTestDlg::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	//rgszNames�Ǹ��ַ������飬cNamesָ������������м����ַ������������1���ַ�����������
	if (cNames != 1)
		return E_NOTIMPL;
	//����ַ�����ShowMessageBox��˵��JavaScript�ڵ�������������ShowMessageBox�������ҾͰ����ⶨ��IDͨ��rgDispId������
	if (wcscmp(rgszNames[0], L"ShowMessageBox") == 0)
	{
		*rgDispId = FUNCTION_ShowMessageBox;
		return S_OK;
	}
	//ͬ������ַ�����GetProcessID��˵��JavaScript�ڵ�������������GetProcessID����
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

//JavaScriptͨ��GetIDsOfNames�õ��ҵĶ���ķ�����ID�󣬻����Invoke��dispIdMember���Ǹղ��Ҹ����������Լ��ⶨ��ID
//wFlagsָ��JavaScript���ҵĶ������ʲô���飡
//�����DISPATCH_METHOD��˵��JavaScript�ڵ����������ķ���������cpp_object.ShowMessageBox();
//�����DISPATCH_PROPERTYGET��˵��JavaScript�ڻ�ȡ�����������ԣ�����var n = cpp_object.num;
//�����DISPATCH_PROPERTYPUT��˵��JavaScript���޸������������ԣ�����cpp_object.num = 10;
//�����DISPATCH_PROPERTYPUTREF��˵��JavaScript��ͨ�������޸�������󣬾�����Ҳ����
//ʾ�����벢û���漰��wFlags�Ͷ������Ե�ʹ�ã���Ҫ���������о����÷���һ����
//pDispParams����JavaScript�����ҵĶ���ķ���ʱ���ݽ����Ĳ�����������һ�����鱣�������в���
//pDispParams->cArgs�����������ж��ٸ�����
//pDispParams->rgvarg���Ǳ����Ų��������飬��ʹ��[]�±������ʣ�ÿ����������VARIANT���ͣ����Ա���������͵�ֵ
//������ʲô������VARIANT::vt���жϣ���������ˣ�VARIANT�ⶫ����Ҷ���
//pVarResult�������Ǹ�JavaScript�ķ���ֵ
//�������ù�
HRESULT STDMETHODCALLTYPE CwebTestDlg::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
	WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	//ͨ��ID�Ҿ�֪��JavaScript������ĸ�����
	if (dispIdMember == FUNCTION_ShowMessageBox)
	{
		//����Ƿ�ֻ��һ������
		if (pDispParams->cArgs != 1)
			return E_NOTIMPL;
		//�����������Ƿ����ַ�������
		if (pDispParams->rgvarg[0].vt != VT_BSTR)
			return E_NOTIMPL;
		//���ĵ���
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

//JavaScript�õ����Ǵ��ݸ�����ָ�����������������ǵĶ�����ʲô�����������QueryInterface��ѯ�����ǡ�����ʲô��������
//����ͨ��riid����������ʲô������ֻ�����ʵ������ǲ���IID_IDispatch�������ǲ���IID_IUnknownʱ�����ǲ��ܿ϶��Ļش���S_OK
//��Ϊ���ǵĶ���̳���IDispatch����IDispatch�ּ̳���IUnknown������ֻʵ�����������ӿڣ�����ֻ���������ش�����ѯ��
HRESULT STDMETHODCALLTYPE CwebTestDlg::QueryInterface(REFIID riid, void **ppvObject)
{
	if (riid == IID_IDispatch || riid == IID_IUnknown)
	{
		//�Եģ�����һ��IDispatch�������Լ�(this)������
		*ppvObject = static_cast<IDispatch*>(this);
		return S_OK;
	}
	else
		return E_NOINTERFACE;
}

//����֪��COM����ʹ�����ü�������������������ڣ����ǵ�CJsCallCppDlg������������ھ��������������������
//�ҵ����������Ҫ��JavaScript���ܣ����Լ���ܣ������Ҳ���ʵ��AddRef()��Release()��������дһЩ��
//��Ҫreturn 1;return 2;return 3;return 4;return 5;������
ULONG STDMETHODCALLTYPE CwebTestDlg::AddRef()
{
	return 1;
}

//ͬ�ϣ�����˵��
//���⻰����Ȼ�����Ҫnew��һ��c++���������Ӹ�JavaScript���ܣ������Ҫʵ��AddRef()��Release()�������ü�������ʱdelete this;
ULONG STDMETHODCALLTYPE CwebTestDlg::Release()
{
	return 1;
}BEGIN_EVENTSINK_MAP(CwebTestDlg, CDHtmlDialog)
	ON_EVENT(CwebTestDlg, IDC_EXPLORER1, 250, CwebTestDlg::BeforeNavigate2Explorer1, VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PBOOL)
	ON_EVENT(CwebTestDlg, IDC_EXPLORER1, 259, CwebTestDlg::DocumentCompleteExplorer1, VTS_DISPATCH VTS_PVARIANT)
	END_EVENTSINK_MAP()


void CwebTestDlg::BeforeNavigate2Explorer1(LPDISPATCH pDisp, VARIANT* URL, VARIANT* Flags, VARIANT* TargetFrameName, VARIANT* PostData, VARIANT* Headers, BOOL* Cancel)
{
	// TODO: �ڴ˴������Ϣ����������
}


//���¼���Ӧ��������WebBrowser����HTML�ĵ���Ϻ󱻵���
//��ʱ����JavaScript��SaveCppObject�����������Լ�(this)��������SaveCppObject�����������󱣴浽ȫ�ֱ���var cpp_object;��
//�Ժ�JavaScript�Ϳ���ͨ��cpp_object�����������C++����ķ�����
//ע�⣬��Ϊ�ҵĳ���ֻ��m_webbrowser.Navigate������һ��HTML������������û�м��VARIANT* URL
//�����ĳ������ж��m_webbrowser.Navigate����Ҫͨ��VARIANT* URL�ж�����¼���Ӧ�Ƕ�Ӧ���Ĵ�Navigate
void CwebTestDlg::DocumentCompleteExplorer1(LPDISPATCH pDisp, VARIANT* URL)
{
	CComQIPtr<IHTMLDocument2> document = m_webbrowser.get_Document();
	CComDispatchDriver script;
	document->get_Script(&script);
	CComVariant var(static_cast<IDispatch*>(this));
	script.Invoke1(L"SaveCppObject", &var);
}
