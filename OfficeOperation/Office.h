#ifndef Office_h__
#define Office_h__
#ifndef  LIBNAME
#undef  LIBNAME
#endif

#define  LIBNAME "OfficeOperation.lib"
#include "vc.h"

#if _MSC_VER < 1300
#include "vc6\inc\OfficeOperationInc.h"
#elif _MSC_VER ==1300
#include "vc7\inc\OfficeOperationInc.h"
#elif _MSC_VER <1400
#include "vc7.1\inc\OfficeOperationInc.h"
#elif _MSC_VER <1500
#include "vc8\inc\OfficeOperationInc.h"
#elif _MSC_VER ==1500
#include "vc9\inc\OfficeOperationInc.h"
#elif _MSC_VER ==1600
#include "vc10\inc\OfficeOperationInc.h"
#else
#error "Unsupported VC++ version"
#endif // Office_h__
