
// webTestDlg.h : 头文件
//

#pragma once

#include "explorer1.h"
// CwebTestDlg 对话框
class CwebTestDlg : public CDHtmlDialog, public IDispatch 
{
// 构造
public:
	CwebTestDlg(CWnd* pParent = NULL);	// 标准构造函数
	int  showRes;
// 对话框数据
	enum { IDD = IDD_WEBTEST_DIALOG, IDH = IDR_HTML_WEBTEST_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

	HRESULT OnButtonOK(IHTMLElement *pElement);
	HRESULT OnButtonCancel(IHTMLElement *pElement);
	bool ProcessHtmlCommand(LPCTSTR lpszURL);
	DWORD GetProcessID();
	
	void ShowMessageBox(const wchar_t *msg);  
	DWORD test1();
private:
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo);  

	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);  

	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);  

	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);  

	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppvObject);  

	virtual ULONG STDMETHODCALLTYPE AddRef();  

	virtual ULONG STDMETHODCALLTYPE Release();  
	CExplorer1 m_webbrowser;
// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	DECLARE_DHTML_EVENT_MAP()
public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	//virtual void OnBeforeNavigate(LPDISPATCH pDisp, LPCTSTR szUrl);
	//void _OnBeforeNavigate2(LPDISPATCH pDisp, VARIANT FAR* URL, VARIANT FAR* Flags, VARIANT FAR* TargetFrameName, VARIANT FAR* PostData, VARIANT FAR* Headers, BOOL FAR* Cancel);

	virtual void OnBeforeNavigate(LPDISPATCH pDisp, LPCTSTR szUrl);
	afx_msg void OnBnClickedButton1();


	DECLARE_EVENTSINK_MAP()
	void BeforeNavigate2Explorer1(LPDISPATCH pDisp, VARIANT* URL, VARIANT* Flags, VARIANT* TargetFrameName, VARIANT* PostData, VARIANT* Headers, BOOL* Cancel);
	void DocumentCompleteExplorer1(LPDISPATCH pDisp, VARIANT* URL);
};
