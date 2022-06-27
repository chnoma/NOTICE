import pandas as pd

# region Constants
SITE_DETAILS = pd.read_excel("./settings/SiteList.xlsx")
SITE_DETAILS.set_index("Station#", inplace=True)


# endregion


def jam(d, return_value=""):  # If string is None or nan, return ""
    o = str(d)
    if o == "None" or o == "nan":
        return return_value
    return o


def jam_int(d, return_value=0):
    try:
        d = int(d)
    except ValueError:
        return return_value
    return d


def code_to_site(site_code, area):
    output = {}
    multi = False
    # THIS SECTION FINDS DUPLICATE VA ID'S AND USES THE PROVIDED NAME TO DECONFLICT THEM AND FIND THE RIGHT ONE
    if isinstance(SITE_DETAILS["District"][site_code], pd.core.series.Series):
        for k, v in enumerate(SITE_DETAILS["District"][site_code].reset_index()):
            if area == SITE_DETAILS["Area"][site_code].reset_index()["Area"][k]:
                multi = True
                output["District"] = jam(SITE_DETAILS["District"][site_code].reset_index()["District"][k])
                output["Area"] = jam(SITE_DETAILS["Area"][site_code].reset_index()["Area"][k])
                output["MyVA VISN"] = jam(SITE_DETAILS["MyVA VISN"][site_code].reset_index()["MyVA VISN"][k])
                output["Area"] = jam(SITE_DETAILS["Area"][site_code].reset_index()["Area"][k])
                output["Site"] = jam(site_code)
                output["Location Code"] = jam(
                    SITE_DETAILS["Location Code"][site_code].reset_index()["Location Code"][k])
                output["Shipping Address"] = jam(
                    SITE_DETAILS["Shipping Address"][site_code].reset_index()["Shipping Address"][k])
                output["Shipping City"] = jam(
                    SITE_DETAILS["Shipping City"][site_code].reset_index()["Shipping City"][k])
                output["Shipping State"] = jam(
                    SITE_DETAILS["Shipping State"][site_code].reset_index()["Shipping State"][k])
                output["Shipping Zip Code"] = jam(
                    SITE_DETAILS["Shipping Zip Code"][site_code].reset_index()["Shipping Zip Code"][k])
                output["E-mail Distribution List for Logistics"] = jam(
                    SITE_DETAILS["E-mail Distribution List for Logistics"][site_code].reset_index()[
                        "E-mail Distribution List for Logistics"][k])
                output["E-mail Distribution List for OIT"] = jam(
                    SITE_DETAILS["E-mail Distribution List for OIT"][site_code].reset_index()[
                        "E-mail Distribution List for OIT"][k])
                output["Delivery POC"] = jam(
                    SITE_DETAILS["Delivery POC"][site_code].reset_index()["Delivery POC"][k])
                output["Delivery POC Cell Phone#"] = jam(
                    SITE_DETAILS["Delivery POC Cell Phone#"][site_code].reset_index()["Delivery POC Cell Phone#"][k])
                output["Delivery POC Email"] = jam(
                    SITE_DETAILS["Delivery POC Email"][site_code].reset_index()["Delivery POC Email"][k])
                output["Delivery POC Phone#"] = jam(
                    SITE_DETAILS["Delivery POC Phone#"][site_code].reset_index()["Delivery POC Phone#"][k])
                output["Alternate POC"] = jam(
                    SITE_DETAILS["Alternate POC"][site_code].reset_index()["Alternate POC"][k])
                output["Alternate POC Cell Phone#"] = jam(
                    SITE_DETAILS["Alternate POC Cell Phone#"][site_code].reset_index()["Alternate POC Cell Phone#"][k])
                output["Alternate POC Email"] = jam(
                    SITE_DETAILS["Alternate POC Email"][site_code].reset_index()["Alternate POC Email"][k])
                output["Alternate POC Phone#"] = jam(
                    SITE_DETAILS["Alternate POC Phone#"][site_code].reset_index()["Alternate POC Phone#"][k])
    if not multi:
        output["District"] = jam(SITE_DETAILS["District"][site_code])
        output["Area"] = jam(SITE_DETAILS["Area"][site_code])
        output["MyVA VISN"] = jam(SITE_DETAILS["MyVA VISN"][site_code])
        output["Area"] = jam(SITE_DETAILS["Area"][site_code])
        output["Site"] = jam(site_code)
        output["Location Code"] = jam(SITE_DETAILS["Location Code"][site_code])
        output["Shipping Address"] = jam(SITE_DETAILS["Shipping Address"][site_code])
        output["Shipping City"] = jam(SITE_DETAILS["Shipping City"][site_code])
        output["Shipping State"] = jam(SITE_DETAILS["Shipping State"][site_code])
        output["Shipping Zip Code"] = jam(SITE_DETAILS["Shipping Zip Code"][site_code])
        output["E-mail Distribution List for Logistics"] = jam(
            SITE_DETAILS["E-mail Distribution List for Logistics"][site_code])
        output["E-mail Distribution List for OIT"] = jam(
            SITE_DETAILS["E-mail Distribution List for OIT"][site_code])
        output["Delivery POC"] = jam(SITE_DETAILS["Delivery POC"][site_code])
        output["Delivery POC Cell Phone#"] = jam(SITE_DETAILS["Delivery POC Cell Phone#"][site_code])
        output["Delivery POC Email"] = jam(SITE_DETAILS["Delivery POC Email"][site_code])
        output["Delivery POC Phone#"] = jam(SITE_DETAILS["Delivery POC Phone#"][site_code])
        output["Alternate POC"] = jam(SITE_DETAILS["Alternate POC"][site_code])
        output["Alternate POC Cell Phone#"] = jam(SITE_DETAILS["Alternate POC Cell Phone#"][site_code])
        output["Alternate POC Email"] = jam(SITE_DETAILS["Alternate POC Email"][site_code])
        output["Alternate POC Phone#"] = jam(SITE_DETAILS["Alternate POC Phone#"][site_code])
    return output
