//////////////////////////////////////////////////////////////////////
/// \file UndocComctl32.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Undocumented comctl32.dll stuff</em>
///
/// Declaration of some comctl32.dll internals that we're using.
///
/// \sa Pager
//////////////////////////////////////////////////////////////////////

#pragma once


/// \brief <em>Retrieves a listview control's color for hot items' text</em>
///
/// This is documented, yet undefined window message to configure automatic scrolling of a pager control.
///
/// \param[in] wParam Specifies the timeout value for the scroll events, in milliseconds.
/// \param[in] lParam The \c LOWORD specifies the number of lines to scroll per timeout. The \c HIWORD
///            specifies the number of pixels per line.
///
/// \return The return value is not used.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/ee663599.aspx">PGM_SETSCROLLINFO</a>
#define PGM_SETSCROLLINFO (PGM_FIRST + 13)