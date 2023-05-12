import win32com.client
import pythoncom

swConst = win32com.client.gencache.EnsureModule(
    "{4687F359-55D0-4CD3-B6CF-2EB42C11F989}", 0, 29, 0
).constants  # sw2015
swCmd = win32com.client.gencache.EnsureModule(
    "{0AC1DE9F-3FBC-4C25-868D-7D4E9139CCE0}", 0, 29, 0
).constants
# # {83A33D31-27C5-11CE-BFD4-00400513BB57}
arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
var1 = win32com.client.VARIANT(
    pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None
)  # (pythoncom.VT_VARIANT | pythoncom.VT_NULL | pythoncom.VT_BYREF, None)
varVals = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, -1)
varBool = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, None)
varNone = win32com.client.VARIANT(pythoncom.VT_EMPTY, None)
Empty = win32com.client.VARIANT(pythoncom.VT_EMPTY, None)
newDrw = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, -1)



def converUnits_XYZ(X, Y, Z, unit):
    # converts various units to metres:
    returnX = convertUnits(X, unit)
    returnY = convertUnits(Y, unit)
    returnZ = convertUnits(Z, unit)
    return returnX, returnY, returnZ


def convertUnits(distance, unit):
    # converts various units to metres:
    if unit == "mm":
        returnDistance = distance / 1000
    if unit == "in":
        returnDistance = distance * 0.0254
    if unit == "ft":
        returnDistance = distance * (0.0254 * 12)
    return returnDistance