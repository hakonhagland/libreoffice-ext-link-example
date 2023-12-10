#! /usr/bin/env python3

import logging
import os
import uno
from unohelper import systemPathToFileUrl, absolutize
from com.sun.star.beans import PropertyValue


def main() -> None:
    logging.basicConfig(level=logging.INFO)
    # Connect to the running instance of LibreOffice or start a new one
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    # Load the document
    cwd = systemPathToFileUrl( os.getcwd() )
    path = '1.fodt'
    file_url = absolutize( cwd, systemPathToFileUrl(path) )

    load_props = []
    #load_props.append(PropertyValue(Name="Hidden", Value=True))  # Open document hidden
    load_props.append(PropertyValue(Name="UpdateDocMode",
                Value=uno.getConstantByName("com.sun.star.document.UpdateDocMode.FULL_UPDATE")))
    load_props = tuple(load_props)

    logging.info("Loading {}".format(file_url))
    doc = desktop.loadComponentFromURL(file_url, "_blank", 0, load_props)

    # Save the document
    #doc.store()
    #doc.dispose()
    logging.info("Done")

if __name__ == "__main__":
    main()
