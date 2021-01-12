# -*- coding: gb2312 -*-

import sys, win32com.client

def main ():

    try:
        visio = win32com.client.Dispatch("Visio.Application")
        #visio.Visible = 0

        dwg = visio.Documents.Open('C:\\Users\\Administrator\\PycharmProjects\\QualityCheckSharepoint\\test\\配置管理流程图-v0.1.vsdx')

        # Used by Visio Shape.BoundingBox method
        intFlags = 0
        visBBoxUprightWH = 0x1

        try:

            vsoShapes = dwg.Pages.Item(1).Shapes # Get shapes for Visio Page-1


            for s in range (len (vsoShapes)):


                # This line works
                print "Index = %s, Shape = %s, Text = %s, Type = %s" % (vsoShapes[s].Index, vsoShapes[s].Name, vsoShapes[s].Text, vsoShapes[s].Type)



                dblLeft =0.0
                dblBottom =0.0
                dblRight = 0.0
                dblTop = 0.0

                # ====== This line will fail with invalid syntax =======
                #vsoShapes.Item(s).BoundingBox(intFlags + visBBoxUprightWH, dblLeft, dblBottom, dblRight, dblTop)

        except Exception, e:
              print "Error", e
        #dwg.Close()
        visio.Quit()

    except Exception, e:
        print "Error opening visio file",e
        visio.Quit()

main()