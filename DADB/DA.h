#pragma once
#include <atlbase.h>
#include <ATLComTime.h>
#include <atldbcli.h>

#include "OpcEnum.h"
#include "opccomn.h"
#include "opcda.h"

HRESULT DA(CLSID& cidOpcServer, CSession& session);