//////////////////////////////////////////////////////////////////////
/// \class Proxy_IOLEDataObjectEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IOLEDataObjectEvents interface</em>
///
/// \if UNICODE
///   \sa TargetOLEDataObject, PagerCtlLibU::_IOLEDataObjectEvents
/// \else
///   \sa TargetOLEDataObject, PagerCtlLibA::_IOLEDataObjectEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IOLEDataObjectEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IOLEDataObjectEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IOLEDataObjectEvents