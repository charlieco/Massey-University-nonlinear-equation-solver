from CoolProp.CoolProp import PropsSI



def enthalpy_pt(f, v):
    try:
        return (PropsSI("H", "P", v[0] * 100000, "T", (v[1] + 273.15), f) / 1000)
    except:
        try:
            return (PropsSI("H", "P", (v[0] + 1) * 100000, "T", (v[1]  + 273.15), f) / 1000)
        except:
            return (PropsSI("H", "P", v[0] * 100000, "T", (v[1]  + 273.15 + 1), f) / 1000)


def enthalpy_ps(f, v):
    try:
        return (PropsSI("H", "P", v[0] * 100000, "S", v[1] * 1000, f) / 1000)
    except:
        try:
            return (PropsSI("H", "P", (v[0] + 1) * 100000, "S", v[1] * 1000, f) / 1000)
        except:
            return (PropsSI("H", "P", (v[0] - 1) * 100000, "S", v[1] * 1000, f) / 1000)


def enthalpy_px(f, v):
    try:
        return (PropsSI("H", "P",v[0] * 100000, "Q", v[1], f) / 1000)
    except:
        try:
            return (PropsSI("H", "P", (v[0] + 1) * 100000, "Q", v[1], f) / 1000)
        except:
            return (PropsSI("H", "P", (v[0] - 1) * 100000, "Q", v[1], f) / 1000)


def entropy_pt(f, v):
    try:
        return (PropsSI("S", "P", v[0] * 100000, "T", (v[1] + 273.15), f) / 1000)
    except:
        try:
            return (PropsSI("S", "P", (v[0]  + 1) * 100000, "T", (v[1] + 273.15), f) / 1000)
        except:
            return (PropsSI("S", "P", v[0]  * 100000, "T", (v[1]+ 273.15 + 1), f) / 1000)


def pressure_tx(f, v):
    try:
        return (PropsSI("P", "T", (v[0] + 273.15), "Q", v[1], f) / 100000)
    except:
        return (PropsSI("P", "T", (v[0] + 273.15 - 1), "Q", v[1], f) / 100000)


def temperature_hp(f, v):
    try:
        return (PropsSI("T", "P", v[1] * 100000, "H", v[0] * 1000, f) - 273.15)
    except:
        return (PropsSI("T", "P", v[1] * 100000, "H", v[0] * 1000, f) - 273.15)

def volume_pt(f, v):
    return 1/PropsSI("D", "P", v[0] * 100000, "T", (v[1]+273.15), f)