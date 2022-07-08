#pragma once
#include "IObjectsCalculator.h"

extern long g_lObjs;
extern long g_lLocks;

class ObjectsCalculator : public IObjectsCalculator
{
protected:
	long          m_lRef;

public:
	ObjectsCalculator();
	~ObjectsCalculator();

public:
	// IUnknown
	STDMETHOD(QueryInterface(REFIID, void**));
	STDMETHOD_(ULONG, AddRef());
	STDMETHOD_(ULONG, Release());

	// IObjectsCalculator
	STDMETHOD(CalculateObjects());

private: 
	void EnumerateRegKey(HKEY hKey) const;
};

class ObjectsCalculatorClassFactory : public IClassFactory
{
protected:
	long       m_lRef;

public:
	ObjectsCalculatorClassFactory();
	~ObjectsCalculatorClassFactory();

	// IUnknown
	STDMETHOD(QueryInterface(REFIID, void**));
	STDMETHOD_(ULONG, AddRef());
	STDMETHOD_(ULONG, Release());

	// IClassFactory
	STDMETHOD(CreateInstance(LPUNKNOWN, REFIID, void**));
	STDMETHOD(LockServer(BOOL));
};