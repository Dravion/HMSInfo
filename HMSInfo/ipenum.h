
#ifndef __IPENUM_H__
#define __IPENUM_H__



class CIPEnum
{
public:
//constructors / Destructors
  CIPEnum();
	virtual ~CIPEnum();

//methods
  BOOL Enumerate();

protected:
  virtual BOOL EnumCallbackFunction(int nAdapter, const in_addr& address) = 0;
  BOOL m_bWinsockInitialied;
};




#endif //__IPENUM_H__