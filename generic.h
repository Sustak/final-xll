#ifndef genericH
#define genericH

#include <ctype.h>
#include <windows.h>
#include "xlcall.h"
#include <math.h>
#include <string.h>

// This is the maximum allowable value for gMaxFuncParms
// it's hard-wired to avoid dynamic memory allocation for
// temporary variables
#define	kMaxSize	32

// function registration data
extern	LPSTR	*gFuncParms;
extern	char	*gModuleName;
extern	int      gFunctionCount, gMaxFuncParms;
extern  char    xloperstrbuf[257];

// public function prototypes
short PStrCmp( const char *str1, const char *str2 );
void PStrCopy( const char *str1, char *str2 );
void PStrCat( char *str1, const char *str2 );
char *NumPStr( int value, char *string );

//==================================================================
// ToPascalString - convert a null-terminated C string to a pascal
//                  string
//
// Arguments
//	string	the C string to be converted in place
// Return Values
//		void
//==================================================================
void ToPascalString( char *string );

// ClipSize is a utility function that will determine the size of a "multi" array
// structure.  It checks to see if the data is organized in columns or rows (giving
// preference to columns), and ignores empty cells at the end of the array.
// It returns the size of the 1D table of valid data.

WORD ClipSize( XLOPER *multi );

void StrToXL( XLOPER& xl, const char* astr, bool allocnewstr=false );
void NumToXL( XLOPER& xl, double anum );
void ErrToXL( XLOPER& xl, int anum );
void CleanXL( XLOPER& xlop );
double XLToNum( const XLOPER& xlop );
void XLToStr( const XLOPER& xlop, char* buf );

unsigned int AllocCounter();

//==================================================================

#endif
