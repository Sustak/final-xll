#
# Generated Makefile - do not edit!
#
# Edit the Makefile in the project folder instead (../Makefile). Each target
# has a -pre and a -post target defined where you can add customized code.
#
# This makefile implements configuration specific macros and targets.


# Environment
MKDIR=mkdir
CP=cp
GREP=grep
NM=nm
CCADMIN=CCadmin
RANLIB=ranlib
CC=gcc.exe
CCC=g++.exe
CXX=g++.exe
FC=gfortran
AS=as.exe

# Macros
CND_PLATFORM=MinGW-Windows
CND_DLIB_EXT=dll
CND_CONF=Debug
CND_DISTDIR=dist
CND_BUILDDIR=build

# Include project Makefile
include Makefile

# Object Directory
OBJECTDIR=${CND_BUILDDIR}/${CND_CONF}/${CND_PLATFORM}

# Object Files
OBJECTFILES= \
	${OBJECTDIR}/generic.o \
	${OBJECTDIR}/finalexcel.o


# C Compiler Flags
CFLAGS=

# CC Compiler Flags
CCFLAGS=-m32 -DBUILDING_DLL=1 -static-libstdc++ -static-libgcc -DWIN32
CXXFLAGS=-m32 -DBUILDING_DLL=1 -static-libstdc++ -static-libgcc -DWIN32

# Fortran Compiler Flags
FFLAGS=

# Assembler Flags
ASFLAGS=

# Link Libraries and Options
LDLIBSOPTIONS=XLCALL32.LIB ../final/dist/Release/MinGW-Windows/libfinal.a ../utils/dist/Release/MinGW-Windows/libutils.a

# Build Targets
.build-conf: ${BUILD_SUBPROJECTS}
	"${MAKE}"  -f nbproject/Makefile-${CND_CONF}.mk ${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/libfinal-xll.${CND_DLIB_EXT}

${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/libfinal-xll.${CND_DLIB_EXT}: XLCALL32.LIB

${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/libfinal-xll.${CND_DLIB_EXT}: ../final/dist/Release/MinGW-Windows/libfinal.a

${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/libfinal-xll.${CND_DLIB_EXT}: ../utils/dist/Release/MinGW-Windows/libutils.a

${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/libfinal-xll.${CND_DLIB_EXT}: ${OBJECTFILES}
	${MKDIR} -p ${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}
	${LINK.cc} -Wl,--add-stdcall-alias --export-all-symbols -shared -o ${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/libfinal-xll.${CND_DLIB_EXT} ${OBJECTFILES} ${LDLIBSOPTIONS} 

${OBJECTDIR}/generic.o: generic.cpp 
	${MKDIR} -p ${OBJECTDIR}
	${RM} $@.d
	$(COMPILE.cc) -g -w  -MMD -MP -MF $@.d -o ${OBJECTDIR}/generic.o generic.cpp

${OBJECTDIR}/finalexcel.o: finalexcel.cpp 
	${MKDIR} -p ${OBJECTDIR}
	${RM} $@.d
	$(COMPILE.cc) -g -w  -MMD -MP -MF $@.d -o ${OBJECTDIR}/finalexcel.o finalexcel.cpp

# Subprojects
.build-subprojects:
	cd ../final && ${MAKE}  -f Makefile CONF=Release
	cd ../utils && ${MAKE}  -f Makefile CONF=Release

# Clean Targets
.clean-conf: ${CLEAN_SUBPROJECTS}
	${RM} -r ${CND_BUILDDIR}/${CND_CONF}
	${RM} ${CND_DISTDIR}/${CND_CONF}/${CND_PLATFORM}/libfinal-xll.${CND_DLIB_EXT}

# Subprojects
.clean-subprojects:
	cd ../final && ${MAKE}  -f Makefile CONF=Release clean
	cd ../utils && ${MAKE}  -f Makefile CONF=Release clean

# Enable dependency checking
.dep.inc: .depcheck-impl

include .dep.inc
