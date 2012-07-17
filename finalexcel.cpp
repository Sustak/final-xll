// Financial Analytics Library (FINAL) - Add-In for Microsoft Excel
// Copyright (c) 2004 - 2012 by Marek Sestak, marek.sestak@gmail.com
//
//    This program is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    This program is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with this program.  If not, see <http://www.gnu.org/licenses/>.

#include "../final/final.h"
#include "generic.h"

//#define		kFunctionCount	5
#define		kMaxFuncParms	27
#define         kMaxResults     256

char	*gModuleName = "\043Financial Analytics Library (FINAL)";

LPSTR functionParms[/*kFunctionCount*/][kMaxFuncParms] =
{
//	function title, argument types, function name, arg names, type (1=func,2=cmd),
//		group name (func wizard), hotkey, help ID, func help string, (repeat) argument help strings

    {" FINAL_DEBUG_AllocStrings",	" R",	" FINAL_DEBUG_AllocStrings",	" ",	" 1",
    	" FINAL DEBUG functions",	" ",
    	" ",
    	" Displays number of strings allocated by FINAL that hasn't yet been freed.",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " " },

    {" FINAL_Test",	" RR",	" FINAL_Test",	" ",	" 1",
    	" FINAL DEBUG functions",	" ",
    	" ",
    	" Displays number of strings allocated by FINAL that hasn't yet been freed.",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " " },

    {" FINAL_Version",	" R",	" FINAL_Version",	" ",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Displays FINAL's build version, date and time.",
    	" ", " ", " ", " ", " ", " ", " ", " ",	" ", " ", " ", " ", " ", " ", " ", " ", " ", " " },

    {" FIL_Version",	" R",	" FIL_Version",	" ",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Version.)\nDisplays FINAL's build version, date and time.",
    	" ", " ", " ", " ", " ", " ", " ", " ",	" ", " ", " ", " ", " ", " ", " ", " ", " ", " " },

    {" FINAL_About",	" R",	" FINAL_About",	" ",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Displays FINAL's copyright & license information.",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " " },


    {" FINAL_All", " RJJBBBJJCJJJBJ", " FINAL_All",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon, NominalAmount, Labels",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Calculates a series of security's parameters:\n"
          "- Market Value\n"
          "- Conventional Yield\n"
          "- Money Market Yield\n"
          "- Duration\n"
          "- Modified Duration\n"
          "- Risk measure\n"
          "- Convexity\n"
          "- Previous Coupon\n"
          "- Next Coupon\n"
          "- Next Coupon Amount\n"
          "- Accrued Interest\n"
          "- Duration Amount\n"
          "- Value of 1bp (BPV)\n"
          "- Security Type\n"
          "- Number of remaining cashflows\n"
          "- Cashflow number 1 date\n"
          "- Cashflow number 1 amount\n"
          "- Additional cashflows up to the maturity of the security",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        "  is amount, at which calculations are performed (optional, default 1,000,000.00)",
        "  are included in the first column (optional, default false)",
        " ", " ", " ", " ", " "
    },

    {" FIL_All", " RJJBBBJJCJJJBJ", " FIL_All",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon, NominalAmount, Labels",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Calculates a series of security's parameters:\n"
    	  "- Conventional Yield",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        "  is amount, at which calculations are performed (optional, default 1,000,000.00)",
        "  are included in the first column (optional, default false)",
        " ", " ", " ", " ", " "
    },

    {" FINAL_Yield", " RJJBBBJJCJJJ", " FINAL_Yield",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns yield for a fixed income security (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_Yield", " RJJBBBJJCJJJ", " FIL_Yield",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Yield.)\nReturns yield for a fixed income security (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_MMktYield", " RJJBBBJJCJJJ", " FINAL_MMktYield",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns money market yield for a fixed income security (bill, bond, cd, cp, strip...)\n"
          "If there's more than one cashflow until the maturity (i.e. the final one) "
          "#VALUE is returned",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_MMktYield", " RJJBBBJJCJJJ", " FIL_MMktYield",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Yield.)\nReturns money market yield for a fixed income security (bill, bond, cd, cp, strip...)\n"
          "If there's more than one cashflow until the maturity (i.e. the final one) "
          "#VALUE is returned",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_Discount", " RJJBBBJJCJJJ", " FINAL_Discount", " Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns discount for a fixed income security (bill, bond, cd, cp, strip...)\n"
          "If there's more than one cashflow until the maturity (i.e. the final one) "
          "#VALUE is returned.",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's price",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_Discount", " RJJBBBJJCJJJ", " FIL_Discount",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Yield.)\nReturns discount for a fixed income security (bill, bond, cd, cp, strip...)\n"
          "If there's more than one cashflow until the maturity (i.e. the final one) "
          "#VALUE is returned.",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's price",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_Price", " RJJBBBJJCJJJ", " FINAL_Price",	" Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Converts fixed income security's conventional yield into price",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield that will be converted to price",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_Price", " RJJBBBJJCJJJ", " FIL_Price",	" Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Price.)\nConverts fixed income security's conventional yield into price",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield that will be converted to price",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_MMktYieldToPrice", " RJJBBBJJCJJJ", " FINAL_MMktYieldToPrice", " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Converts fixed income security's conventional yield into price.\n"
          "If there's more than one cashflow until the maturity (i.e. the final one) "
          "#VALUE is returned",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is money market yield that will be converted to price",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_MMktYieldToPrice", " RJJBBBJJCJJJ", " FIL_MMktYieldToPrice", " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Price.)\nConverts fixed income security's conventional yield into price\n"
          "If there's more than one cashflow until the maturity (i.e. the final one) "
          "#VALUE is returned",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is money market yield that will be converted to price",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_DiscountToPrice", " RJJBBBJJCJJJ", " FINAL_DiscountToPrice", " Settlement, Maturity, Coupon, Discount, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Converts fixed income security's discount into price.\n"
          "If there's more than one cashflow until the maturity (i.e. the final one) "
          "#VALUE is returned",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's discount that will be converted to price",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_DiscountToPrice", " RJJBBBJJCJJJ", " FIL_DiscountToPrice", " Settlement, Maturity, Coupon, Discount, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Price.)\nConverts fixed income security's discount into price.\n"
          "If there's more than one cashflow until the maturity (i.e. the final one) "
          "#VALUE is returned",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's discount that will be converted to price",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_Duration", " RJJBBBJJCJJJ", " FINAL_Duration",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns fixed income security's duration (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_Duration", " RJJBBBJJCJJJ", " FIL_Duration",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Duration.)\nReturns fixed income security's duration (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_DurationY", " RJJBBBJJCJJJ", " FINAL_DurationY",	" Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns fixed income security's duration (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_DurationY", " RJJBBBJJCJJJ", " FIL_DurationY", " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Duration.)\nReturns fixed income security's duration (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_MDuration", " RJJBBBJJCJJJ", " FINAL_MDuration",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns fixed income security's modified duration (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_MDuration", " RJJBBBJJCJJJ", " FIL_MDuration",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_MDuration.)\nReturns fixed income security's modified duration (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_MDurationY", " RJJBBBJJCJJJ", " FINAL_MDurationY",	" Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns fixed income security's modified duration (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_MDurationY", " RJJBBBJJCJJJ", " FIL_MDurationY", " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Duration.)\nReturns fixed income security's modified duration (bill, bond, cd, cp, strip...)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_Risk", " RJJBBBJJCJJJ", " FINAL_Risk",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns fixed income security's risk measure",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_Risk", " RJJBBBJJCJJJ", " FIL_Risk",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Risk.)\nReturns fixed income security's risk measure",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_RiskY", " RJJBBBJJCJJJ", " FINAL_RiskY",	" Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns fixed income security's risk measure",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_RiskY", " RJJBBBJJCJJJ", " FIL_RiskY", " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Duration.)\nReturns fixed income security's risk measure",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_Convexity", " RJJBBBJJCJJJ", " FINAL_Convexity",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns fixed income security's convexity",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_Convexity", " RJJBBBJJCJJJ", " FIL_Convexity",	" Settlement, Maturity, Coupon, Price, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Convexity.)\nReturns fixed income security's convexity",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is price at which all calculations are performed",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_ConvexityY", " RJJBBBJJCJJJ", " FINAL_ConvexityY",	" Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns fixed income security's convexity",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_ConvexityY", " RJJBBBJJCJJJ", " FIL_ConvexityY", " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_Duration.)\nReturns fixed income security's convexity",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_AccruedInterest", " RJJBJJCJJJ", " FINAL_AccruedInterest", " Settlement, Maturity, Coupon, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns fixed income security's accrued interest",
    	"  is settlement date at which accrued interest is calculated",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_AccruedInterest", " RJJBJJCJJJ", " FIL_AccruedInterest", " Settlement, Maturity, Coupon, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)", " ",
    	" ",
        " (Obsolete version. Use FINAL_AccruedInterest.)\n"
    	  "Returns fixed income security's accrued interest",
    	"  is settlement date at which accrued interest is calculated",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_CouponPeriod", " RJJBJJCJJJ", " FINAL_CouponPeriod", " Settlement, Maturity, Coupon, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns information about security's coupon:\n"
          "- date of previous coupon (or a date from which the coupon accrues)\n"
          "- date of next coupon\n"
          "- coupon amount that will be payed out on the next coupon date",
    	"  is settlement date",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_CouponPeriod", " RJJBJJCJJJ", " FIL_CouponPeriod", " Settlement, Maturity, Coupon, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon",	" 1",
    	" FIL functions (obsolete)", " ",
    	" ",
        " (Obsolete version. Use FINAL_AccruedInterest.)\n"
          "Returns information about income security's coupon:\n"
          "- date of previous coupon (or a date from which the coupon accrues)\n"
          "- date of next coupon\n"
          "- coupon amount that will be payed out on the next coupon date",
    	"  is settlement date",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_YearFrac", " RJJJ", " FINAL_YearFrac",	" Start_Date, End_Date, Basis",	" 1",
    	" FINAL functions",	" ",
    	" ",
    	" Calculates fraction of year between Start_Date and End_Date according to a daycount basis defined in Basis ",
    	"  is the beginning of the period",
    	"  is the end of the period",
    	"  is the type of day count convention that is used to calculate the fraction of year",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_YearFrac", " RJJJ", " FIL_YearFrac",	" Start_Date, End_Date, Basis",	" 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_YearFrac.)\nCalculates fraction of year between Start_Date and End_Date according to a daycount basis defined in Basis ",
    	"  is the beginning of the period",
    	"  is the end of the period",
    	"  is the type of day count convention that is used to calculate the fraction of year",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_IsLeap", " RJ", " FINAL_IsLeap", " Year", " 1",
    	" FINAL functions",	" ",
    	" ",
    	" Return true for leap years",
    	"  for which it is to determine whether the year is leap or not",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_IsLeap", " RJ", " FIL_IsLeap", " Year", " 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_IsLeap.)\nReturn true for leap years",
    	"  for which it is to determine whether the year is leap or not",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_EOMonth", " RJJ", " FINAL_EOMonth", " Start, Months", " 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns end of month given number of months after or before start date.\n"
          "If Months is set to zero, end of the current month is returned.",
    	"  is the start date",
        "  is the number of months after (>0) or before (<0) the start date",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_EOMonth", " RJJ", " FIL_EOMonth", " Start, Months", " 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_EOMonth.)\nReturns end of month given number of months after or before start date.\n"
          "If Months is set to zero, end of the current month is returned.",
    	"  is the start date",
        "  is the number of months after (>0) or before (<0) the start date",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_EDate", " RJJ", " FINAL_EDate", " Start, Months", " 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns date given number of months after or before start date.\n"
          "If Months is set to zero, start date is returned.",
    	"  is the start date",
        "  is the number of months after (>0) or before (<0) the start date",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_EDate", " RJJ", " FIL_EDate", " Start, Months", " 1",
    	" FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_EDate.) Returns date given number of months after or before start date.\n"
          "If Months is set to zero, start date is returned.",
    	"  is the start date",
        "  is the number of months after (>0) or before (<0) the start date",
    	" ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_CouponNum", " RJJJJ", " FINAL_CouponNum", " Settlement, Maturity, Frequency, Basis", " 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns number of coupons that will be paid out between date settlement and maturity date. "
          "(Coupons paid out on settlement date are excluded, coupons on maturity date included)",
    	"  is the date, from which coupons are counted",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_CouponNum", " RJJJJ", " FIL_CouponNum", " Settlement, Maturity, Frequency, Basis", " 1",
    	" FIL functions (obsolete)", " ",
    	" ",
    	" (Obsolete version. Use FINAL_CouponNum.)\nReturns number of coupons that will be paid out between date settlement and maturity date. "
          "(Coupons paid out on settlement date are excluded, coupons on maturity date included)",
    	"  is the date, from which coupons are counted",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_PrevCouponDate", " RJJJJ", " FINAL_PrevCouponDate", " Settlement, Maturity, Frequency, Basis", " 1",
    	" FINAL functions",	" ",
    	" ",
    	" Returns coupon date preceding the settlement date."
          "(If the settlement date falls on a coupon date, this settlement date is returned.)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_PrevCouponDate", " RJJJJ", " FIL_PrevCouponDate", " Settlement, Maturity, Frequency, Basis", " 1",
    	" FIL functions (obsolete)", " ",
    	" ",
    	" (Obsolete version. Use FINAL_PrevCouponDate.)\n"
          "Returns coupon date preceding the settlement date."
          "(If the settlement date falls on a coupon date, this settlement date is returned.)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_NextCouponDate", " RJJJJ", " FINAL_NextCouponDate", " Settlement, Maturity, Frequency, Basis", " 1",
   	" FINAL functions",	" ",
    	" ",
    	" Returns coupon date following the settlement date.\n"
          "(If the settlement date falls on a coupon date, next coupon date is returned)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_NextCouponDate", " RJJJJ", " FIL_NextCouponDate", " Settlement, Maturity, Frequency, Basis", " 1",
    	" FIL functions (obsolete)", " ",
    	" ",
    	" (Obsolete version. Use FINAL_NextCouponDate.)\n"
          "Returns coupon date following the settlement date.\n"
          "(If the settlement date falls on a coupon date, next coupon date is returned)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_CouponDays", " RJJJJ", " FINAL_CouponDays", " Settlement, Maturity, Frequency, Basis", " 1",
   	" FINAL functions",	" ",
    	" ",
    	" Returns number of days in the coupon period containing the settlement date.\n"
          "(If the settlement date falls on a coupon date, coupon period from the settlement to the next coupon date is considered)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_CouponDays", " RJJJJ", " FIL_CouponDays", " Settlement, Maturity, Frequency, Basis", " 1",
    	" FIL functions (obsolete)", " ",
    	" ",
    	" (Obsolete version. Use FINAL_CouponDays.)\n"
          "Returns number of days in the coupon period containing the settlement date.\n"
          "(If the settlement date falls on a coupon date, coupon period from the settlement to the next coupon date is considered)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_CouponDaysBS", " RJJJJ", " FINAL_CouponDaysBS", " Settlement, Maturity, Frequency, Basis", " 1",
   	" FINAL functions",	" ",
    	" ",
    	" Returns number of days from the begining of the current coupon period to the settlement date.\n"
          "(If the settlement date falls on a coupon date, zero is returned)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_CouponDaysBS", " RJJJJ", " FIL_CouponDaysBS", " Settlement, Maturity, Frequency, Basis", " 1",
    	" FIL functions (obsolete)", " ",
    	" ",
    	" (Obsolete version. Use FINAL_CouponDaysBS.)\n"
    	  "Returns number of days from the begining of the current coupon period to the settlement date.\n"
          "(If the settlement date falls on a coupon date, zero is returned)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_CouponDaysNC", " RJJJJ", " FINAL_CouponDaysNC", " Settlement, Maturity, Frequency, Basis", " 1",
   	" FINAL functions",	" ",
    	" ",
    	" Returns number of days from the settlement date to the next coupon date.\n"
          "(If the settlement date falls on a coupon date, number of days in the whole coupon period is returned)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_CouponDaysNC", " RJJJJ", " FIL_CouponDaysNC", " Settlement, Maturity, Frequency, Basis", " 1",
    	" FIL functions (obsolete)", " ",
    	" ",
    	" (Obsolete version. Use FINAL_CouponDaysNC.)\n"
    	  "Returns number of days from the settlement date to the next coupon date.\n"
          "(If the settlement date falls on a coupon date, number of days in the whole coupon period is returned)",
    	"  is the settlement date",
        "  is the security's maturity date",
    	"  is the security's coupon frequency",
        "  defines the security's coupon day count convention",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FINAL_ASWSpread", " RJJBBBJJCJJJJRRRJJJ", " FINAL_ASWSpread",
        " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon, "
          "BaseDate, Deposits, Futures, Swap, SwapFrequency, FixedBasis, FloatBasis",
        " 1", " FINAL functions",	" ",
    	" ",
    	" Calculates security's asset swap spread (ASW spread)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        "  defines settlement date for deposit rates and interest rate swaps",
        "  defines deposit rates\n(array of two columns and at least one row,\n"
          "first column contains description of the term (1W,1M,3M,1Y etc.)\n"
          "corresponding deposit rate is in the second column)",
        "  defines 3M futures rates\n(array of two columns, first column contains delivery date of the contract\n"
          "second column contains corresponding yield of the 3M Libor/Euribor contract\noptional",
        "  defines swap rates\n(array of two columns, first column contains description of the term (1Y, 2Y, 3Y etc.)\n"
          "corresponding swap rate is in the second column;\n"
          "optional)",
        "  is swap's fixed coupon payment frequency (optional, default 1)",
        "  is swap's fixed coupon day count convention (optional, default US-30/360)",
        "  is swap's floating coupon day count convention (optional, default US-30/360)"
    },

    {" FIL_ASW", " RJJBBBJJCJJJJRRRJJJ", " FIL_ASWSpread",
        " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon, "
          "BaseDate, Deposits, Futures, Swap, SwapFrequency, FixedBasis, FloatBasis",
        " 1", " FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_ASWSpread.)\n"
          "Calculates security's asset swap spread (ASW spread)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        "  defines settlement date for deposit rates and interest rate swaps",
        "  defines deposit rates\n(array of two columns and at least one row,\n"
          "first column contains description of the term (1W,1M,3M,1Y etc.)\n"
          "corresponding deposit rate is in the second column)",
        "  defines 3M futures rates\n(array of two columns, first column contains delivery date of the contract\n"
          "second column contains corresponding yield of the 3M Libor/Euribor contract\noptional",
        "  defines swap rates\n(array of two columns, first column contains description of the term (1Y, 2Y, 3Y etc.)\n"
          "corresponding swap rate is in the second column;\n"
          "optional)",
        "  is swap's fixed coupon payment frequency (optional, default 1)",
        "  is swap's fixed coupon day count convention (optional, default US-30/360)",
        "  is swap's floating coupon day count convention (optional, default US-30/360)"
     },

     {" FINAL_ZSpread", " RJJBBBJJCJJJJRRRJJJ", " FINAL_ZSpread",
        " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon, "
          "BaseDate, Deposits, Futures, Swap, SwapFrequency, FixedBasis, FloatBasis",
        " 1", " FINAL functions",	" ",
    	" ",
    	" Calculates security's Z-spread (a spread to a zero coupon libor curve)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        "  defines settlement date for deposit rates and interest rate swaps",
        "  defines deposit rates\n(array of two columns and at least one row,\n"
          "first column contains description of the term (1W,1M,3M,1Y etc.)\n"
          "corresponding deposit rate is in the second column)",
        "  defines 3M futures rates\n(array of two columns, first column contains delivery date of the contract\n"
          "second column contains corresponding yield of the 3M Libor/Euribor contract\noptional",
        "  defines swap rates\n(array of two columns, first column contains description of the term (1Y, 2Y, 3Y etc.)\n"
          "corresponding swap rate is in the second column;\n"
          "optional)",
        "  is swap's fixed coupon payment frequency (optional, default 1)",
        "  is swap's fixed coupon day count convention (optional, default US-30/360)",
        "  is swap's floating coupon day count convention (optional, default US-30/360)"
    },

    {" FIL_Z", " RJJBBBJJCJJJJRRRJJJ", " FIL_ZSpread",
        " Settlement, Maturity, Coupon, Yield, Redemption, Frequency, Basis, Model, Issued, IntAccrDate, FirstCoupon, "
          "BaseDate, Deposits, Futures, Swap, SwapFrequency, FixedBasis, FloatBasis",
        " 1", " FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_ASWSpread.)\n"
          "Calculates security's asset swap spread (ASW spread)",
    	"  is settlement date at which all calculations are performed",
    	"  is security's maturity date",
        "  is security's annual coupon rate (e.g. for 4% pass 0.04)",
    	"  is security's yield to maturity",
    	"  is security's redemption value (optional, default 100)",
        "  is security's coupon frequency (optional, default 1)",
        "  is security's accrued interest day count basis (optional, default 0=US 30/360)",
        "  defines security's calculation model (optional, default is generic security)",
        "  defines when the security was issued (optional, default is theoretical coupon date preceeding settlement date)",
        "  defines when the security's first coupon starts to accrue (optional, default is start of the coupon period containing issue date)",
        "  defines when the security pays out first coupon (optional, default is end of the coupon period containing issue date)",
        "  defines settlement date for deposit rates and interest rate swaps",
        "  defines deposit rates\n(array of two columns and at least one row,\n"
          "first column contains description of the term (1W,1M,3M,1Y etc.)\n"
          "corresponding deposit rate is in the second column)",
        "  defines 3M futures rates\n(array of two columns, first column contains delivery date of the contract\n"
          "second column contains corresponding yield of the 3M Libor/Euribor contract\noptional",
        "  defines swap rates\n(array of two columns, first column contains description of the term (1Y, 2Y, 3Y etc.)\n"
          "corresponding swap rate is in the second column;\n"
          "optional)",
        "  is swap's fixed coupon payment frequency (optional, default 1)",
        "  is swap's fixed coupon day count convention (optional, default US-30/360)",
        "  is swap's floating coupon day count convention (optional, default US-30/360)"
     },

    {" FINAL_ZeroCouponLiborCurve", " RJRRRJJJ", " FINAL_ZeroCouponLiborCurve",
        " BaseDate, Deposits, Futures, Swap, SwapFrequency, FixedBasis, FloatBasis",
        " 1", " FINAL functions",	" ",
    	" ",
    	" Converts deposit rates, money market futures and swap rates into a zero coupon 'libor' curve",
        "  defines settlement date for deposit rates and interest rate swaps",
        "  defines deposit rates\n(array of two columns and at least one row,\n"
          "first column contains description of the term (1W,1M,3M,1Y etc.)\n"
          "corresponding deposit rate is in the second column)",
        "  defines 3M futures rates\n(array of two columns, first column contains delivery date of the contract\n"
          "second column contains corresponding yield of the 3M Libor/Euribor contract\noptional",
        "  defines swap rates\n(array of two columns, first column contains description of the term (1Y, 2Y, 3Y etc.)\n"
          "corresponding swap rate is in the second column;\n"
          "optional)",
        "  is swap's fixed coupon payment frequency (optional, default 1)",
        "  is swap's fixed coupon day count convention (optional, default US-30/360)",
        "  is swap's floating coupon day count convention (optional, default US-30/360)",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

    {" FIL_SwapCurveToZeroCpn", " RJRRRJJJ", " FIL_SwapCurveToZeroCpn",
        " BaseDate, Deposits, Futures, Swap, SwapFrequency, FixedBasis, FloatBasis",
        " 1", " FIL functions (obsolete)",	" ",
    	" ",
    	" (Obsolete version. Use FINAL_ZeroCouponLiborCurve.)\n"
          "Converts deposit rates, money market futures and swap rates into a zero coupon 'libor' curve",
        "  defines settlement date for deposit rates and interest rate swaps",
        "  defines deposit rates\n(array of two columns and at least one row,\n"
          "first column contains description of the term (1W,1M,3M,1Y etc.)\n"
          "corresponding deposit rate is in the second column)",
        "  defines 3M futures rates\n(array of two columns, first column contains delivery date of the contract\n"
          "second column contains corresponding yield of the 3M Libor/Euribor contract\noptional",
        "  defines swap rates\n(array of two columns, first column contains description of the term (1Y, 2Y, 3Y etc.)\n"
          "corresponding swap rate is in the second column;\n"
          "optional)",
        "  is swap's fixed coupon payment frequency (optional, default 1)",
        "  is swap's fixed coupon day count convention (optional, default US-30/360)",
        "  is swap's floating coupon day count convention (optional, default US-30/360)",
        " ", " ", " ", " ", " ", " ", " ", " ", " ", " ", " "
    },

  };


int	gFunctionCount = sizeof( functionParms )/sizeof( functionParms[0] ); // kFunctionCount;
int	gMaxFuncParms = kMaxFuncParms;
LPSTR   *gFuncParms = &functionParms[0][0];

char    xloperstrbuf[257];
XLOPER  result;
XLOPER  resultArray[kMaxResults];
// for results of type 'string':
// if they point to a string other than to xloperstrbuf, they were dynamically
// allocated and should be therefore freed by us too!
unsigned int nresultsUsed = 0;
// when your functions intends to return an array, you should set how many
// elements of the array are used (needed for instance in xlAutoFree!)


const char* MSG_ABOUT_FINAL = "Financial analytics library (FINAL), (c) 2002-2005 Marek �est�k (marek.sestak@gmail.com), licensed under GNU General Public License 2 (see http://www.gnu.org/copyleft/gpl.html for details)";
const char* MSG_UNABLE_TO_BUILD_SECURITY = "Unable to build security from the parameters passed.";
const char* MSG_INVALID_MATURITY_DATE = "Invalid maturity date.";
const char* MSG_UNABLE_TO_BUILD_SWAP_CURVE = "Unable to build a swap curve";
const char* MSG_UNABLE_TO_BUILD_SWAP_CURVE_NO_LIBORS = "Unable to build a swap curve: at least one libor rate has to be provided.";

using namespace final;
using namespace utils;

extern "C" __declspec(dllexport) void WINAPI xlAutoFree( LPXLOPER pxFree )
{
  for( unsigned int i=0; i<nresultsUsed && i<kMaxResults; i++ ) {
    CleanXL( resultArray[i] );
  }

  #ifdef _DEBUG
  TString msg( AllocCounter() );
  MessageBox( NULL, msg.c_str(), "xlAutoFree called, AllocCounter says:", MB_OK );
  #endif
}

// ---------------------------------------------------------------------------

TSecurity* BuildSecurity( int settlement, int maturity, double coupon,
    double redemption, int frequency, int basis, char far *securitymodel,
    int issued, int interestaccrual, int firstcoupon )
{
  TSecurity *sec;
  try {
    if( settlement==0 || maturity==0 )
      throw TExFINAL( MSG_UNABLE_TO_BUILD_SECURITY );

    if( redemption==0.0 )
      redemption = 100.0;

    if( frequency==0 )
      frequency = 1;

    if( issued==0)
      issued = PrevCouponDate( TDate(settlement), TDate(maturity), frequency, basis ).Serial();

    if( interestaccrual==0 )
      interestaccrual = -1;

    if( firstcoupon==0 )
      firstcoupon = -1;

    TString ticker;
    if( securitymodel )
      ticker = TString( securitymodel );

    sec = GetSecurity( ticker, coupon, maturity, issued, frequency, basis,
                       interestaccrual, firstcoupon, redemption );
  }
  catch(...) {
    throw TExFINAL( MSG_UNABLE_TO_BUILD_SECURITY );
  }
  return sec;
}

extern "C" __declspec(dllexport) __stdcall LPXLOPER FINAL_Version()
{
  static XLOPER result;
  StrToXL( result, FINAL );
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall LPXLOPER FIL_Version()
{
  static XLOPER result;
  StrToXL( result, FINAL );
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall LPXLOPER FINAL_About()
{
  static XLOPER result;
  StrToXL( result, MSG_ABOUT_FINAL );
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall LPXLOPER FINAL_DEBUG_AllocStrings()
{
  static XLOPER result;
  NumToXL( result, AllocCounter() );
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_All( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon, double nominalamount, int includelabels )
    /* RJJBBBJJCJJJBJ */
{
  #ifdef _DEBUG
  MessageBox( NULL, "FINAL_All called", "Trace", MB_OK );
  #endif

  XLOPER arg;
  arg.xltype = xltypeBool;
  arg.val.boolean = 0;

//  const int e = Excel4( xlfVolatile, 0, 1, &arg );
//  if( e!=xlretSuccess ) {
//    TString ret(e);
//    MessageBox( NULL, "Failed to make FINAL_ALL not volatile", ret.c_str(), MB_OK );
//  }

  TSecurity *sec = NULL;
  try {

    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

//    if( nominalamount==0.0 )
//      nominalamount = 1000000.0;
// uprimne, nepamatuji si, zda tudle defaultni hodnotu nekde nepotrebuji...

    result.xltype = xltypeMulti;
    result.val.array.lparray = resultArray;

    if( includelabels ) {
      result.xltype |= xlbitDLLFree;
      // since we also return an array of strings, we'd better clean up
      // the allocated mess and free what we had previously allocated
    }

    XLOPER *currValue = resultArray;
    TDate settle( settlement );
    floating mv, yld;
    bool mvfailed = false;

    // --------
    if( includelabels ) StrToXL( *(currValue++), "Market Value", true );
    try { NumToXL( *currValue, mv = sec->MarketValue( nominalamount, settle, price ) ); }
    catch (...) { ErrToXL( *currValue, xlerrNA ); mvfailed=true; }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Conventional Yield", true );
    try { NumToXL( *currValue, yld = sec->ConvYield( settle, price ) ); }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Money Market Yield", true );
    try { NumToXL( *currValue, sec->MMktYield( settle, price ) ); }
    catch( ... ) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Duration", true );
    try { NumToXL( *currValue, sec->DurationY( settle, yld ) ); }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Modified Duration", true );
    try { NumToXL( *currValue, sec->MDurationY( settle, yld ) ); }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Risk", true );
    try { NumToXL( *currValue, sec->RiskY( settle, yld ) ); }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Convexity", true );
    try { NumToXL( *currValue, sec->ConvexityY( settle, yld ) ); }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Previous Coupon", true );
    try {
      if( sec->HasCoupon() )
        NumToXL( *currValue, sec->PrevCouponDate( settle ).Serial() );
      else
        ErrToXL( *currValue, xlerrNA );
    }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Next Coupon", true );
    try {
      if( sec->HasCoupon() )
        NumToXL( *currValue, sec->NextCouponDate( settle ).Serial() );
      else
        ErrToXL( *currValue, xlerrNA );
    }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Next Coupon Amount", true );
    try {
      if( sec->HasCoupon() )
        NumToXL( *currValue, sec->NextCouponSize( settle ) );
      else
        NumToXL( *currValue, 0.0 );
    }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Accrued Interest", true );
    try {
      if( sec->HasCoupon() )
        NumToXL( *currValue, sec->AccruedInterest( settle ) );
      else
        NumToXL( *currValue, 0.0 );
    }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Duration Amount", true );
    if( mvfailed ) ErrToXL( *(currValue++), xlerrNA );
    else           NumToXL( *(currValue++), mv );
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Value of 1bp (BPV)", true );
    try { NumToXL( *currValue, yld = sec->BpvAmountY( nominalamount, settle, yld ) ); }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Security Type", true );
    StrToXL( *(currValue++), sec->SecurityType(), true );

    // ---------------------------
    // cashflows...

    double nextCFAmt;
    int    nCF;
    TDate  nextCFDate;
    try {
      nCF = sec->NextCashflow( settle, nextCFDate, nextCFAmt );
    }
    catch( ... ) {
      nCF = 0;
    }
    // --------
    if( includelabels ) StrToXL( *(currValue++), "Number of remaining cashflows", true );
    NumToXL( *(currValue++), nCF );
    // --------
    int i = 1;
    char label[256];
    while( nCF>0 && (currValue-resultArray)<kMaxResults ) {
      if( includelabels ) {
        sprintf( label, "CF #%d date", i );
        StrToXL( *(currValue++), label, true );
      }
      NumToXL( *(currValue++), nextCFDate.Serial() );
      if( includelabels ) {
        sprintf( label, "CF #%d amount", i );
        StrToXL( *(currValue++), label, true );
      }
      NumToXL( *(currValue++), nextCFAmt * nominalamount / 100.0 );

      try {
        nCF = sec->NextCashflow( nextCFDate+1, nextCFDate, nextCFAmt );
      }
      catch( ... ) {
        nCF = 0;
      }
      i++;
    }


    // ---------------------------
    // default format:
    // if labels are included: 2 columns by X rows
    // without labels:         X columns by 1 row

    if( includelabels ) {
      result.val.array.columns = 2;
      nresultsUsed = currValue - resultArray;
      result.val.array.rows = nresultsUsed / ( includelabels ? 2 : 1 );
    }
    else {
      result.val.array.columns = currValue - resultArray;
      result.val.array.rows = 1;
      nresultsUsed = 0;
    }
  }
  catch( ... ) {
    ErrToXL( result, xlerrValue );
  }

  if( sec )
    delete sec;
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_All( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon, double nominalamount, int includelabels )
    /* RJJBBBJJCJJJBJ */
{
  return FIL_All( settlement, maturity, coupon,
    price, redemption, frequency, basis,
    securitymodel, issued, interestaccrual,
    firstcoupon, nominalamount, includelabels );
}

// ----------------------------------------------------------------------------
// Yield

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_Yield( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->ConvYield( TDate(settlement), price ) );
  }
  catch( ... ) {
    ErrToXL( result, xlerrValue );
  }
  if( sec )
    delete sec;
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_Yield( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_Yield( settlement, maturity, coupon, price, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// MmktYield

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_MMktYield( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->MMktYield( TDate(settlement), price ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_MMktYield( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_MMktYield( settlement, maturity, coupon, price, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// Discount

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_Discount( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->Discount( TDate(settlement), price ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;

 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_Discount( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_Discount( settlement, maturity, coupon, price, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// Price

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_Price( int settlement, int maturity, double coupon,
    double aytm, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->ConvPrice( TDate(settlement), aytm ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_Price( int settlement, int maturity, double coupon,
    double aytm, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_Price( settlement, maturity, coupon, aytm, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// MmktYieldToPrice

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_MMktYieldToPrice( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->MMktPrice( TDate(settlement), yield ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_MMktYieldToPrice( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_MMktYieldToPrice( settlement, maturity, coupon, yield, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// Discount

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_DiscountToPrice( int settlement, int maturity, double coupon,
    double discount, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->DiscountToPrice( TDate(settlement), discount ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_DiscountToPrice( int settlement, int maturity, double coupon,
    double discount, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_DiscountToPrice( settlement, maturity, coupon, discount, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// Duration

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_Duration( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->Duration( TDate(settlement), price ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_Duration( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_Duration( settlement, maturity, coupon, price, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// DurationY

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_DurationY( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->DurationY( TDate(settlement), yield ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_DurationY( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_DurationY( settlement, maturity, coupon, yield, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// Modified Duration

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_MDuration( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->MDuration( TDate(settlement), price ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_MDuration( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_MDuration( settlement, maturity, coupon, price, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// ModifiedDurationY

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_MDurationY( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->MDurationY( TDate(settlement), yield ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_MDurationY( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_MDurationY( settlement, maturity, coupon, yield, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// Risk

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_Risk( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->Risk( TDate(settlement), price ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_Risk( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_Risk( settlement, maturity, coupon, price, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// Risk

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_RiskY( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->RiskY( TDate(settlement), yield ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_RiskY( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_RiskY( settlement, maturity, coupon, yield, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// Convexity

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_Convexity( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->Convexity( TDate(settlement), price ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_Convexity( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_Convexity( settlement, maturity, coupon, price, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// Convexity

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_ConvexityY( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->ConvexityY( TDate(settlement), yield ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_ConvexityY( int settlement, int maturity, double coupon,
    double yield, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon )
    /* RJJBBBJJCJJJ */
{
  return FINAL_Convexity( settlement, maturity, coupon, yield, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------
// AccruedInterest

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_AccruedInterest( int settlement, int maturity, double coupon,
    int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual, int firstcoupon )
    /* RJJBJJCJJJ */
{
 TSecurity *sec = NULL;
 try {
    sec = BuildSecurity( settlement, maturity, coupon, 100.0, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, sec->AccruedInterest( TDate(settlement) ) );
 }
 catch( ... ) {
   ErrToXL( result, xlerrValue );
 }
 if( sec )
   delete sec;
 return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_AccruedInterest( int settlement, int maturity, double coupon,
    int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual, int firstcoupon )
    /* RJJBJJCJJJ */
{
  return FINAL_AccruedInterest( settlement, maturity, coupon,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_YearFrac( int from, int to, int basis )
{
  try {
    NumToXL( result, final::YearFrac( from, to, basis ) );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_YearFrac( int from, int to, int basis )
{
  return FINAL_YearFrac( from, to, basis );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_IsLeap( int year )
{
  try {
    NumToXL( result, IsLeap( year ) );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_IsLeap( int year )
{
  return FINAL_IsLeap( year );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_EOMonth( int start, int months )
{
  try {
    NumToXL( result, EOMonth( TDate(start), months ).Serial() );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_EOMonth( int start, int months )
{
  return FINAL_EOMonth( start, months );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_EDate( int start, int months )
{
  try {
    NumToXL( result, EDate( TDate(start), months ).Serial() );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_EDate( int start, int months )
{
  return FINAL_EDate( start, months );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_CouponNum( int settlement, int maturity, int frequency, int basis )
{
  try {
    NumToXL( result, CouponNum( settlement, maturity, frequency, basis ) );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_CouponNum( int settlement, int maturity, int frequency, int basis )
{
  return FINAL_CouponNum( settlement, maturity, frequency, basis );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_PrevCouponDate( int settlement, int maturity, int frequency, int basis )
{
  try {
    NumToXL( result, PrevCouponDate( settlement, maturity, frequency, basis ).Serial() );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_PrevCouponDate( int settlement, int maturity, int frequency, int basis )
{
  return FINAL_PrevCouponDate( settlement, maturity, frequency, basis );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_NextCouponDate( int settlement, int maturity, int frequency, int basis )
{
  try {
    NumToXL( result, NextCouponDate( settlement, maturity, frequency, basis ).Serial() );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_NextCouponDate( int settlement, int maturity, int frequency, int basis )
{
  return FINAL_NextCouponDate( settlement, maturity, frequency, basis );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_CouponDays( int settlement, int maturity, int frequency, int basis )
{
  try {
    NumToXL( result, CouponDays( settlement, maturity, frequency, basis ) );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_CouponDays( int settlement, int maturity, int frequency, int basis )
{
  return FINAL_CouponDays( settlement, maturity, frequency, basis );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_CouponDaysBS( int settlement, int maturity, int frequency, int basis )
{
  try {
    NumToXL( result, CouponDaysBS( settlement, maturity, frequency, basis ) );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_CouponDaysBS( int settlement, int maturity, int frequency, int basis )
{
  return FINAL_CouponDaysBS( settlement, maturity, frequency, basis );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_CouponDaysNC( int settlement, int maturity, int frequency, int basis )
{
  try {
    NumToXL( result, CouponDaysNC( settlement, maturity, frequency, basis ) );
  }
  catch( ... ) {
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_CouponDaysNC( int settlement, int maturity, int frequency, int basis )
{
  return FINAL_CouponDaysNC( settlement, maturity, frequency, basis );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_CouponPeriod( int settlement, int maturity, double coupon,
    int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual, int firstcoupon )
    /* RJJBJJCJJJ */
{
  TSecurity *sec = NULL;
  try {

    sec = BuildSecurity( settlement, maturity, coupon, 100.0, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    result.xltype = xltypeMulti;
    result.val.array.lparray = resultArray;

    XLOPER *currValue = resultArray;
    TDate settle( settlement );

    try {
      if( sec->HasCoupon() )
        NumToXL( *currValue, sec->PrevCouponDate( settle ).Serial() );
      else
        ErrToXL( *currValue, xlerrNA );
    }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    try {
      if( sec->HasCoupon() )
        NumToXL( *currValue, sec->NextCouponDate( settle ).Serial() );
      else
        ErrToXL( *currValue, xlerrNA );
    }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;
    // --------
    try {
      if( sec->HasCoupon() )
        NumToXL( *currValue, sec->NextCouponSize( settle ) );
      else
        NumToXL( *currValue, 0.0 );
    }
    catch (...) { ErrToXL( *currValue, xlerrNA ); }
    currValue++;

    // ---------------------------
    result.val.array.columns = currValue - resultArray;
    nresultsUsed = 0;
    result.val.array.rows = 1;
  }
  catch( ... ) {
    ErrToXL( result, xlerrValue );
  }

  if( sec )
    delete sec;
  return (LPXLOPER) &result;
}



extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_CouponPeriod( int settlement, int maturity, double coupon,
    int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual, int firstcoupon )
    /* RJJBJJCJJJ */
{
  return FINAL_CouponPeriod( settlement, maturity, coupon, frequency, basis,
    securitymodel, issued, interestaccrual, firstcoupon );
}


// ----------------------------------------------------------------------------

TSwapCurve* BuildSwapCurve( int basedate,
  LPXLOPER pxldeposits, LPXLOPER pxlfutures, LPXLOPER pxlswaps,
  int swapfixfq, int fixedbasis, int floatbasis )
{
//  MessageBox( NULL, "entering buildswapcurve", "", MB_OK );
  struct TLocalVars {
    XLOPER deposits;
    XLOPER futures;
    XLOPER swaps;
    TLocalVars() {
      deposits.val.array.lparray = 0;
      futures.val.array.lparray = 0;
      swaps.val.array.lparray = 0; }
    ~TLocalVars() {
//      MessageBox( NULL, "~TLocalVars", "", MB_OK );
      if( deposits.val.array.lparray ) { Excel4( xlFree, 0, 1, (LPXLOPER) &deposits ); }
      if( futures.val.array.lparray ) { Excel4( xlFree, 0, 1, (LPXLOPER) &futures ); }
      if( swaps.val.array.lparray ) { Excel4( xlFree, 0, 1, (LPXLOPER) &swaps ); }
    }
  } vars;

  XLOPER xlint;
  xlint.xltype = xltypeInt;
  xlint.val.w = xltypeMulti;

//  MessageBox( NULL, "coerce deposits", "", MB_OK );
  if( pxldeposits )
    if( xlretUncalced == Excel4( xlCoerce, (LPXLOPER) &(vars.deposits), 2, pxldeposits, &xlint ) ) {
//      MessageBox( NULL, "coerce deposits failed", "", MB_OK );
      return NULL;
    }

//  MessageBox( NULL, "coerce futures", "", MB_OK );
  if( pxlfutures )
    if( xlretUncalced == Excel4( xlCoerce, (LPXLOPER) &(vars.futures), 2, pxlfutures, &xlint ) ) {
//      MessageBox( NULL, "coerce futures failed", "", MB_OK );
      return NULL;
    }

//  MessageBox( NULL, "coerce swaps", "", MB_OK );
  if( pxlswaps )
    if( xlretUncalced == Excel4( xlCoerce, (LPXLOPER) &(vars.swaps), 2, pxlswaps, &xlint ) ) {
//      MessageBox( NULL, "coerce swaps failed", "", MB_OK );
      return NULL;
    }

  if( swapfixfq==0 )
    swapfixfq = 1;

  if( basedate==0 )
    throw TExFINAL( MSG_UNABLE_TO_BUILD_SWAP_CURVE_NO_LIBORS );

  TTermRateVector libors, swaps;
  TYields futures;

  int i;
  int cols, rows;
  LPXLOPER arr;
  char buf[256];
  int valuedate;
  double rate;
  TTermRate r;

  cols = vars.deposits.val.array.columns;
  rows = vars.deposits.val.array.rows;
  arr  = vars.deposits.val.array.lparray;
  if( cols < 2 || rows < 1 )
    throw TExFINAL( MSG_UNABLE_TO_BUILD_SWAP_CURVE );
  for( i=0; i<rows; i++ ) {
    try {
      XLToStr( arr[i*cols], buf );
      rate = XLToNum( arr[i*cols+1] );
      r = GetTermRate( buf, rate );
      libors.push_back( r );
    } catch( ... ) { }
  }
  if( libors.size() == 0 )
    throw TExFINAL( MSG_UNABLE_TO_BUILD_SWAP_CURVE_NO_LIBORS );

  cols = vars.futures.val.array.columns;
  rows = vars.futures.val.array.rows;
  arr  = vars.futures.val.array.lparray;
  if( cols >= 2 ) {
    for( i=0; i<rows; i++ ) {
      try {
        valuedate = (int) XLToNum( arr[i*cols] );
        rate = XLToNum( arr[i*cols+1] );
        futures[ valuedate ] = rate;
      } catch( ... ) {}
    }
  }

  cols = vars.swaps.val.array.columns;
  rows = vars.swaps.val.array.rows;
  arr  = vars.swaps.val.array.lparray;
  for( i=0; i<rows; i++ ) {
    try {
      XLToStr( arr[i*cols], buf );
      rate = XLToNum( arr[i*cols+1] );
      r = GetTermRate( buf, rate );
      swaps.push_back( r );
    } catch( ... ) {}
  }

  return new TSwapCurve( TDate(basedate), libors[0].rate, libors, futures, swaps,
      fixedbasis, floatbasis, swapfixfq );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_ASWSpread( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon,
    int basedate, LPXLOPER deposits, LPXLOPER futures, LPXLOPER swaps,
    int swapfixfq, int fixedbasis, int floatbasis )
    // RJJBBBJJCJJJJRRRJJJ
{
  struct TLocalVars {
    TSwapCurve *curve;
    TSecurity *sec;
    TLocalVars() { curve = NULL; sec = NULL; }
    ~TLocalVars() { if( curve ) delete curve;
                    if( sec ) delete sec; }
  } vars;

  try {
    #ifdef _DEBUG
    MessageBox( NULL, "Building curve", "FINAL_ASWSpread", MB_OK );
    #endif

    vars.curve = BuildSwapCurve( basedate, deposits, futures, swaps, swapfixfq, fixedbasis, floatbasis );
    if( !vars.curve )
      return NULL;

    #ifdef _DEBUG
    MessageBox( NULL, "Building security", "", MB_OK );
    #endif

    vars.sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );
    NumToXL( result, vars.curve->ASWSpread( *vars.sec, TDate(settlement), price ) );
  }
  catch( TExFINAL& ex ) {
    #ifdef _DEBUG
    MessageBox( NULL, ex.Message.c_str(), "", MB_OK );
    #endif
    ErrToXL( result, xlerrNum );
  }
  catch( ... ) {
    #ifdef _DEBUG
    MessageBox( NULL, "exception", "", MB_OK );
    #endif
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_ASW( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon,
    int basedate, LPXLOPER deposits, LPXLOPER futures, LPXLOPER swaps,
    int swapfixfq, int fixedbasis, int floatbasis )
{
  return FINAL_ASWSpread( settlement, maturity, coupon, price, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon,
    basedate, deposits, futures, swaps, swapfixfq, fixedbasis, floatbasis );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_ZSpread( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon,
    int basedate, LPXLOPER deposits, LPXLOPER futures, LPXLOPER swaps,
    int swapfixfq, int fixedbasis, int floatbasis )
    // RJJBBBJJCJJJJRRRJJJ
{
  struct TLocalVars {
    TSwapCurve *curve;
    TSecurity *sec;
    TLocalVars() { curve = NULL; sec = NULL; }
    ~TLocalVars() { if( curve ) delete curve;
                    if( sec ) delete sec; }
  } vars;

  try {
    #ifdef _DEBUG
    MessageBox( NULL, "Building curve", "FINAL_ASWSpread", MB_OK );
    #endif

    vars.curve = BuildSwapCurve( basedate, deposits, futures, swaps, swapfixfq, fixedbasis, floatbasis );
    if( !vars.curve )
      return NULL;

   #ifdef _DEBUG
    // delete!!! ------------------------
    if( TDate(15,2,2007).Serial()==maturity ) {
    TYields dfs, yields;
    vars.curve->DFs( dfs );
    vars.curve->Yields( yields );

    result.xltype = xltypeMulti;
    result.val.array.lparray = resultArray;
    TYields::iterator dfiter, dfiterend, ylditer, ylditerend;

    ylditer = yields.begin();
    ylditerend = yields.end();
    dfiter = dfs.begin();
    dfiterend = dfs.end();

    int valuedt;

    while( ylditer!=ylditerend || dfiter!=dfiterend ) {
      valuedt = ylditer->first.Serial();

      TString msg = TDate(valuedt).DateString() + ", " +TString( dfiter->second ) + ", " + TString( ylditer->second );
      MessageBox( NULL, msg.c_str(), "FINAL_ZSpread Debug", MB_OK );

      dfiter++;
      ylditer++;
    }
    }
    // delete!!! ------------------------
    #endif


    #ifdef _DEBUG
    MessageBox( NULL, "Building security", "", MB_OK );
    #endif

    vars.sec = BuildSecurity( settlement, maturity, coupon, redemption, frequency,
      basis, securitymodel, issued, interestaccrual, firstcoupon );

    NumToXL( result, vars.curve->ZSpread( *vars.sec, TDate(settlement), price ) );
  }
  catch( TExFINAL& ex ) {
    #ifdef _DEBUG
    MessageBox( NULL, ex.Message.c_str(), "", MB_OK );
    #endif
    ErrToXL( result, xlerrNum );
  }
  catch( ... ) {
    #ifdef _DEBUG
    MessageBox( NULL, "exception", "", MB_OK );
    #endif
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_Z( int settlement, int maturity, double coupon,
    double price, double redemption, int frequency, int basis,
    char far *securitymodel, int issued, int interestaccrual,
    int firstcoupon,
    int basedate, LPXLOPER deposits, LPXLOPER futures, LPXLOPER swaps,
    int swapfixfq, int fixedbasis, int floatbasis )
{
  return FINAL_ZSpread( settlement, maturity, coupon, price, redemption,
    frequency, basis, securitymodel, issued, interestaccrual, firstcoupon,
    basedate, deposits, futures, swaps, swapfixfq, fixedbasis, floatbasis );
}

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_ZeroCouponLiborCurve(
    int basedate, LPXLOPER deposits, LPXLOPER futures, LPXLOPER swaps,
    int swapfixfq, int fixedbasis, int floatbasis )
    // RJRRRJJJ
{
  struct TLocalVars {
    TSwapCurve *curve;
    TLocalVars() { curve = NULL; }
    ~TLocalVars() { if( curve ) delete curve; }
  } vars;

  try {
    vars.curve = BuildSwapCurve( basedate, deposits, futures, swaps, swapfixfq, fixedbasis, floatbasis );
    if( !vars.curve ) {
      ErrToXL( result, xlerrValue );
      return (LPXLOPER) &result;
    }

    TYields dfs, yields;
    vars.curve->DFs( dfs );
    vars.curve->Yields( yields );

    result.xltype = xltypeMulti;
    result.val.array.lparray = resultArray;
    XLOPER *currValue = resultArray;
    TYields::iterator dfiter, dfiterend, ylditer, ylditerend;

    ylditer = yields.begin();
    ylditerend = yields.end();
    dfiter = dfs.begin();
    dfiterend = dfs.end();

    int valuedt;

    while( ylditer!=ylditerend || dfiter!=dfiterend ) {
      valuedt = ylditer->first.Serial();
      if( valuedt != dfiter->first.Serial() )
        ErrToXL( result, xlerrNA );
      NumToXL( *(currValue++), valuedt );
      NumToXL( *(currValue++), dfiter->second );
      NumToXL( *(currValue++), ylditer->second );
      dfiter++;
      ylditer++;
    }

    result.val.array.columns = 3;
    nresultsUsed = 0; // we're not returning strings, therefore we can set it to zero
    result.val.array.rows = ( currValue - resultArray ) / 3;
  }
  catch( TExFINAL& ex ) {
    #ifdef _DEBUG
    MessageBox( NULL, ex.Message.c_str(), "", MB_OK );
    #endif
    ErrToXL( result, xlerrNum );
  }
  catch( ... ) {
    #ifdef _DEBUG
    MessageBox( NULL, "exception", "", MB_OK );
    #endif
    ErrToXL( result, xlerrNum );
  }
  return (LPXLOPER) &result;
}

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FIL_SwapCurveToZeroCpn(
    int basedate, LPXLOPER deposits, LPXLOPER futures, LPXLOPER swaps,
    int swapfixfq, int fixedbasis, int floatbasis )
    // RJRRRJJJ
{
  return FINAL_ZeroCouponLiborCurve( basedate, deposits, futures, swaps,
    swapfixfq, fixedbasis, floatbasis );
}

// ----------------------------------------------------------------------------

/*
extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_Select( LPXLOPER pxldesc, char far *select ) {

  struct TLocalVars {
    XLOPER desc;

    TLocalVars() {
      desc.val.array.lparray = 0;
    }
    ~TLocalVars() {
      if( desc.val.array.lparray ) { Excel4( xlFree, 0, 1, (LPXLOPER) &desc ); }
    }
  } vars;


  XLOPER xlint;
  xlint.xltype = xltypeInt;
  xlint.val.w = xltypeMulti;

  if( pxldesc )
    if( xlretUncalced == Excel4( xlCoerce, (LPXLOPER) &(vars.desc), 2, pxldesc, &xlint ) ) {
      ErrToXL( result, xlerrNA );
      return (LPXLOPER) &result;
    }

  int desc_mat;
  int desc_

} */

// ----------------------------------------------------------------------------

extern "C" __declspec(dllexport) __stdcall
LPXLOPER FINAL_Test( LPXLOPER param ) {
  XLOPER val;
  XLOPER xlint;
  xlint.xltype = xltypeInt;
  xlint.val.w = xltypeMulti;

  TString str(param->xltype);
  MessageBox( NULL, str.c_str(), "test", MB_OK );
  if( param->xltype == xltypeNil ) {
    MessageBox( NULL, "null passed", "test", MB_OK );
    return (LPXLOPER) &result;
  }

  if( xlretUncalced == Excel4( xlCoerce, (LPXLOPER) &val, 2, (LPXLOPER) param,
						(LPXLOPER) &xlint ) ) {
    MessageBox( NULL, "Source cell not calculated yet", "Debug", MB_OK );
    return 0;
  }

  char buf[256], out[1024];
  out[0] = 0;

  LPXLOPER ptr = val.val.array.lparray;

  if( ptr->xltype == xltypeStr ) {
    if( ptr->val.str == NULL )
      MessageBox( NULL, "NULL!", "FINAL_Test called", MB_OK );
    else {
      for( int i = 0; i<val.val.array.rows; i++ ) {
        for( int j = 0; j<val.val.array.columns; j++ ) {
          memmove( buf, &(ptr->val.str[1]), ptr->val.str[0] );
          buf[ ptr->val.str[0] ] = 0;
          strcat( out, buf );
          strcat( out, "|" );
          ptr++;
        }
      }
      MessageBox( NULL, out, "FINAL_Test called", MB_OK );
    }
  }
  else
    MessageBox( NULL, "Unknown type", "FINAL_Test called", MB_OK );

  StrToXL( result, "Done" );
  return (LPXLOPER) &result;
}
