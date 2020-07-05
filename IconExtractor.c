#include <windows.h>
#include <io.h>
#include <fcntl.h>
#include <stdio.h>
#include <shlwapi.h>
#include "IconExtractor.h"
#include "resource.h"

// g (Global optimization), s (Favor small code), y (No frame pointers).
#pragma optimize("gsy", on)
#pragma comment(linker, "/FILEALIGN:0x200")
#pragma comment(linker, "/MERGE:.rdata=.data")
#pragma comment(linker, "/MERGE:.text=.data")
#pragma comment(linker, "/MERGE:.reloc=.data")
#pragma comment(linker, "/SECTION:.text, EWR /IGNORE:4078")
#pragma comment(linker, "/OPT:NOWIN98")		// Make section alignment really small.
#define WIN32_LEAN_AND_MEAN

#define WIDTHBYTES(bits)      ((((bits) + 31)>>5)<<2)

LRESULT			CALLBACK Main(HWND hDlg,UINT uMsg,WPARAM wParam,LPARAM lParam);
DWORD			ExtractIconFromExe(LPCSTR, LPCSTR, UINT);
BOOL			AddResource(LPCTSTR,LPTSTR);
DWORD			WriteIconToICOFile( LPICONRESOURCE, LPCTSTR );
BOOL			AdjustIconImagePointers( LPICONIMAGE );
LPSTR			FindDIBBits(LPSTR);
WORD			DIBNumColors(LPSTR);
WORD			PaletteSize(LPSTR);
DWORD			BytesPerLine( LPBITMAPINFOHEADER );
DWORD			WriteICOHeader( HANDLE, UINT );
DWORD			CalculateImageOffset( LPICONRESOURCE, UINT );
BOOL			CALLBACK EnumResLangProc(HMODULE, LPCTSTR, LPCTSTR, WORD, LONG);
BOOL			CALLBACK EnumResNameProc(HMODULE, LPCTSTR, LPTSTR, LONG);

HINSTANCE		hInst;
HWND			hDlg;
char			szFileName[MAX_PATH]= "";
char			iconfilename[MAX_PATH]= "";
char			*szBuffer;
ULONG			GroupIconID[5];
char			GroupIconName[5];
UINT			Counter;
UINT			NameCounter;

int WINAPI WinMain(HINSTANCE hinstCurrent,HINSTANCE hinstPrevious,LPSTR lpszCmdLine,int nCmdShow){
	hInst= hinstCurrent;
	DialogBox(hinstCurrent,MAKEINTRESOURCE(IDD_MAIN),NULL,(DLGPROC)Main);
	return 0;
}

LRESULT CALLBACK Main(HWND hDlgMain,UINT uMsg,WPARAM wParam,LPARAM lParam){
	HDC				hdc;
	PAINTSTRUCT		ps;
	OPENFILENAME	ofn;
	HANDLE			hLoadFile= NULL;

	hDlg= hDlgMain;

	switch(uMsg){
		case WM_INITDIALOG:
			SendMessageA(hDlg,WM_SETICON,ICON_SMALL, (LPARAM) LoadIcon(hInst,MAKEINTRESOURCE(IDI_ICON)));
			SendMessageA(hDlg,WM_SETICON, ICON_BIG,(LPARAM) LoadIcon(hInst,MAKEINTRESOURCE(IDI_ICON)));
		break;
		case WM_COMMAND:
			if( LOWORD(wParam)==IDC_LOADFILE ){
				if( HIWORD(wParam)==BN_CLICKED ){
					ZeroMemory(&ofn,sizeof(ofn));
					szFileName[0]= '\0';
					ofn.lStructSize = sizeof(OPENFILENAME);
					ofn.hwndOwner=hDlg;
					ofn.lpstrFilter= "Executables (*.exe)\0*.exe\0";
					ofn.lpstrFile= szFileName;
					ofn.nMaxFile= MAX_PATH;
					ofn.Flags= OFN_EXPLORER|OFN_PATHMUSTEXIST;

					if( GetOpenFileName(&ofn)==TRUE ){
						if( hLoadFile!=NULL ){
							CloseHandle(hLoadFile);
							hLoadFile= NULL;
						}
						hLoadFile= CreateFile(szFileName, GENERIC_READ,FILE_SHARE_READ,NULL,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
						if( hLoadFile==INVALID_HANDLE_VALUE ){
							MessageBox(hDlg,"No File or could not open the selected file.","Error accessing file",MB_OK);
						}else{
							SetDlgItemText(hDlg, IDC_FILENAME, szFileName);
						}
						CloseHandle(hLoadFile);
					}
				}
			}else if( LOWORD(wParam)==IDC_EXTRACT ){
				if( HIWORD(wParam)==BN_CLICKED ){
					ExtractIconFromExe(szFileName, "1.ico", 0);
				}
			}else if( LOWORD(wParam)==IDC_EXIT ){
				if( HIWORD(wParam)==BN_CLICKED ){
					EndDialog(hDlg,wParam);
					DestroyWindow(hDlg);
				}
			}
		break;
		case WM_PAINT:
			hdc= BeginPaint(hDlg, &ps);
			InvalidateRect(hDlg, NULL, TRUE);
			EndPaint (hDlg, &ps);
		return 0;
		case WM_CLOSE:			
			EndDialog(hDlg,wParam);
			DestroyWindow(hDlg);
		break;
		case WM_DESTROY:
			PostQuitMessage(0);
		break;
	}
	return 0;
}

DWORD ExtractIconFromExe(LPCSTR SourceEXE, LPCSTR TargetICON, UINT IconIndex){
	LPICONRESOURCE      lpIR= NULL;
	HINSTANCE        	hLibrary;
	DWORD               ret;
	HRSRC				hRsrc= NULL;
    HGLOBAL				hGlobal= NULL;
    LPMEMICONDIR		lpIcon= NULL;
	UINT				i;

	//ASSERT( IconIndex<5 );

    // Load the EXE - NOTE: must be a 32bit EXE for this to work
    if( (hLibrary= LoadLibraryEx( SourceEXE, NULL, LOAD_LIBRARY_AS_DATAFILE ))==NULL ){
		return GetLastError();
	}

	Counter= 0;
    NameCounter= 0;

 	ret= EnumResourceNames(hLibrary, RT_GROUP_ICON, EnumResNameProc, NULL);

	if( !ret ){
		ret= GetLastError();
		FreeLibrary(hLibrary);
		return ret;
	}

	// Find the group icon resource
	// Finding the resource with the ID first
	if( Counter>0)
		hRsrc= FindResource(hLibrary,MAKEINTRESOURCE(GroupIconID[IconIndex]), RT_GROUP_ICON);

	if( hRsrc==NULL){// Searching the icon with names if any
		hRsrc= FindResource(hLibrary, GroupIconName[IconIndex], RT_GROUP_ICON);
	}

	if( hRsrc==NULL ){
		ret= GetLastError();
		FreeLibrary(hLibrary);
		return ret;
	}

    if( (hGlobal= LoadResource( hLibrary, hRsrc ))==NULL ){
        FreeLibrary( hLibrary );
		return GetLastError();
    }
    if( (lpIcon= (LPMEMICONDIR)LockResource(hGlobal))==NULL ){
        FreeLibrary( hLibrary );
		return GetLastError();
    }
 
	// Allocate enough memory for the images
    if( (lpIR= (LPICONRESOURCE) malloc( sizeof(ICONRESOURCE) + ((lpIcon->idCount-1) * sizeof(ICONIMAGE)) ))==NULL ){
		FreeLibrary( hLibrary );
		return GetLastError();
	}
 

	// Fill in local struct members
    lpIR->nNumImages= lpIcon->idCount;
   

    // Loop through the images
    for( i = 0; i < lpIR->nNumImages; i++ ){
		// Get the individual image
        if( (hRsrc= FindResource( hLibrary, MAKEINTRESOURCE(lpIcon->idEntries[i].nID), RT_ICON )) == NULL ){
			FreeLibrary( hLibrary );
	  		free(lpIR);
   			return GetLastError();
        }
        if( (hGlobal= LoadResource( hLibrary, hRsrc ))==NULL ){
			FreeLibrary( hLibrary );
	  		free(lpIR);
    		return GetLastError();
        }
		// Store a copy of the resource locally
        lpIR->IconImages[i].dwNumBytes= SizeofResource( hLibrary, hRsrc );

	    lpIR->IconImages[i].lpBits =(LPBYTE) malloc( lpIR->IconImages[i].dwNumBytes );
		if( lpIR->IconImages[i].lpBits==NULL ){
			free(lpIR);
			return GetLastError();
		}

        memcpy( lpIR->IconImages[i].lpBits, LockResource( hGlobal ), lpIR->IconImages[i].dwNumBytes );

        // Adjust internal pointers
        if( ! AdjustIconImagePointers( &(lpIR->IconImages[i]) ) ){
			FreeLibrary( hLibrary );
			free(lpIR);
    		return GetLastError();
        }
	}

    FreeLibrary( hLibrary );
	ret= WriteIconToICOFile(lpIR,TargetICON);	

	if(ret){
		free(lpIR);
		return ret;
	}

	free(lpIR);
	return NO_ERROR;
}

DWORD WriteIconToICOFile( LPICONRESOURCE lpIR, LPCTSTR szFileName){
    HANDLE			hFile;
    UINT			i;
    DWORD			dwBytesWritten, dwTemp;
	ICONDIRENTRY	ide;

    // open the file
    if( (hFile= CreateFile( szFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL ))==INVALID_HANDLE_VALUE ){
        return GetLastError();
    }
    // Write the header
    if( WriteICOHeader( hFile, lpIR->nNumImages ) ){
        CloseHandle( hFile );
        return GetLastError();
    }
    // Write the ICONDIRENTRY's
    for(i = 0; i < lpIR->nNumImages; i++ ){
        // Convert internal format to ICONDIRENTRY
        ide.bWidth= lpIR->IconImages[i].Width;
        ide.bHeight= lpIR->IconImages[i].Height;
        ide.bReserved= 0;
        ide.wPlanes= lpIR->IconImages[i].lpbi->bmiHeader.biPlanes;
        ide.wBitCount= lpIR->IconImages[i].lpbi->bmiHeader.biBitCount;
        if( (ide.wPlanes * ide.wBitCount)>=8 ){
            ide.bColorCount = 0;
        }else{
            ide.bColorCount= 1 << (ide.wPlanes * ide.wBitCount);
		}
        ide.dwBytesInRes= lpIR->IconImages[i].dwNumBytes;
        ide.dwImageOffset= CalculateImageOffset( lpIR, i );
 
		// Write the ICONDIRENTRY out to disk
        if( !WriteFile( hFile, &ide, sizeof( ICONDIRENTRY ), &dwBytesWritten, NULL ) ){
			return GetLastError();
		}
        // Did we write a full ICONDIRENTRY ?
        if( dwBytesWritten!=sizeof( ICONDIRENTRY ) ){
			return GetLastError();
		}
    }
    // Write the image bits for each image
    for( i=0; i < lpIR->nNumImages; i++ ){
        dwTemp= lpIR->IconImages[i].lpbi->bmiHeader.biSizeImage;

        // Set the sizeimage member to zero
        lpIR->IconImages[i].lpbi->bmiHeader.biSizeImage = 0;
        // Write the image bits to file
        if( ! WriteFile( hFile, lpIR->IconImages[i].lpBits, lpIR->IconImages[i].dwNumBytes, &dwBytesWritten, NULL ) )
			return GetLastError();
        if( dwBytesWritten != lpIR->IconImages[i].dwNumBytes )
			return GetLastError();
        // set it back
        lpIR->IconImages[i].lpbi->bmiHeader.biSizeImage = dwTemp;
    }
    CloseHandle( hFile );
    return NO_ERROR;
}

BOOL CALLBACK EnumResLangProc(HMODULE hModule, LPCTSTR lpszType, LPCTSTR lpszName, WORD wIDLanguage, LONG lParam){
	HANDLE hUpdate;
	hUpdate= BeginUpdateResource(szFileName, FALSE);
	if( hUpdate==NULL ){
		return TRUE;
	}
	if( UpdateResource(hUpdate, lpszType, lpszName, wIDLanguage, NULL, 0)==FALSE ){
		return TRUE;
	}
	EndUpdateResource(hUpdate, FALSE);
	return TRUE;
}
/*
BOOL CALLBACK EnumResNameProc(HMODULE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG lParam){
	EnumResourceLanguages(hModule, lpszType, lpszName, (ENUMRESLANGPROC)EnumResLangProc, NULL);
	return TRUE;
}
*/

BOOL CALLBACK EnumResNameProc(HMODULE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG lParam){
    return AddResource(lpszType,lpszName);
}

BOOL AdjustIconImagePointers( LPICONIMAGE lpImage ){
    // Sanity check
    if( lpImage==NULL ){
        return FALSE;
	}
    // BITMAPINFO is at beginning of bits
    lpImage->lpbi= (LPBITMAPINFO)lpImage->lpBits;
    // Width - simple enough
    lpImage->Width= lpImage->lpbi->bmiHeader.biWidth;
    // Icons are stored in funky format where height is doubled - account for it
    lpImage->Height= (lpImage->lpbi->bmiHeader.biHeight)/2;
    // How many colors?
    lpImage->Colors= lpImage->lpbi->bmiHeader.biPlanes * lpImage->lpbi->bmiHeader.biBitCount;
    // XOR bits follow the header and color table
    lpImage->lpXOR= (PBYTE)FindDIBBits((LPSTR)lpImage->lpbi);
    // AND bits follow the XOR bits
    lpImage->lpAND= lpImage->lpXOR + (lpImage->Height*BytesPerLine((LPBITMAPINFOHEADER)(lpImage->lpbi)));
    return TRUE;
}

LPSTR FindDIBBits( LPSTR lpbi ){
   return ( lpbi + *(LPDWORD)lpbi + PaletteSize( lpbi ) );
}


WORD PaletteSize( LPSTR lpbi ){
    return ( DIBNumColors( lpbi ) * sizeof( RGBQUAD ) );
}

DWORD BytesPerLine( LPBITMAPINFOHEADER lpBMIH ){
    return WIDTHBYTES(lpBMIH->biWidth * lpBMIH->biPlanes * lpBMIH->biBitCount);
}

WORD DIBNumColors( LPSTR lpbi ){
    WORD	wBitCount;
    DWORD	dwClrUsed;

    dwClrUsed= ((LPBITMAPINFOHEADER) lpbi)->biClrUsed;

    if( dwClrUsed ){
        return (WORD) dwClrUsed;
	}

    wBitCount= ((LPBITMAPINFOHEADER) lpbi)->biBitCount;

    switch( wBitCount ){
        case 1: return 2;
        case 4: return 16;
        case 8:	return 256;
        default:return 0;
    }
    return 0;
}

BOOL AddResource(LPCTSTR lpszType,LPTSTR lpszName){
	if( Counter>=5 || NameCounter>=5 )
		return TRUE;

	if( lpszType==RT_GROUP_ICON ){
		if( (ULONG)lpszName<65536 ){
			GroupIconID[Counter]=  (ULONG)lpszName ;
			Counter += 1;
		}else{
			GroupIconName[NameCounter]= lpszName ;
			NameCounter++;

		}

	}
	return TRUE;
}

DWORD CalculateImageOffset( LPICONRESOURCE lpIR, UINT nIndex ){
    DWORD	dwSize;
    UINT    i;

    // Calculate the ICO header size
    dwSize= 3 * sizeof(WORD);
    // Add the ICONDIRENTRY's
    dwSize += lpIR->nNumImages * sizeof(ICONDIRENTRY);
    // Add the sizes of the previous images
    for( i=0;i<nIndex;i++ )
        dwSize += lpIR->IconImages[i].dwNumBytes;
    // we're there - return the number
    return dwSize;
}

DWORD WriteICOHeader( HANDLE hFile, UINT nNumEntries ){
    WORD    Output;
    DWORD	dwBytesWritten;

    // Write 'reserved' WORD
    Output= 0;
    if( ! WriteFile( hFile, &Output, sizeof( WORD ), &dwBytesWritten, NULL ) )
        return GetLastError();
    // Did we write a WORD?
    if( dwBytesWritten!=sizeof( WORD ) )
        return GetLastError();
    // Write 'type' WORD (1)
    Output = 1;
    if( ! WriteFile( hFile, &Output, sizeof( WORD ), &dwBytesWritten, NULL ) )
        return GetLastError();
    // Did we write a WORD?
    if( dwBytesWritten != sizeof( WORD ) )
        return GetLastError();
    // Write Number of Entries
    Output = (WORD)nNumEntries;
    if( ! WriteFile( hFile, &Output, sizeof( WORD ), &dwBytesWritten, NULL ) )
        return GetLastError();
    // Did we write a WORD?
    if( dwBytesWritten != sizeof( WORD ) )
        return GetLastError();
    return NO_ERROR;
}