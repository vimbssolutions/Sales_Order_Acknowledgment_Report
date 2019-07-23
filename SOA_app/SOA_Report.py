from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.lib.pagesizes import letter, landscape ,inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.platypus import SimpleDocTemplate, Spacer, Paragraph
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import BaseDocTemplate,PageTemplate, Frame
from reportlab.platypus.tables import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.colors import Color
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.graphics import shapes
from reportlab.graphics.charts.axes import XCategoryAxis,YValueAxis
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import Paragraph, SimpleDocTemplate
from reportlab.lib.units import mm
from reportlab.lib.units import inch, cm
from reportlab.platypus import Paragraph, Frame, KeepInFrame
from reportlab.lib.colors import Color, toColor
from reportlab.lib.colors import pink, black, red, blue, green,white, green,gray
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus.frames import Frame
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Paragraph
from reportlab.lib.enums import TA_CENTER
import io
from reportlab.pdfgen import canvas
import os
from reportlab.platypus import BaseDocTemplate, Frame, Paragraph, NextPageTemplate, PageBreak, PageTemplate
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
from reportlab.platypus import BaseDocTemplate,Frame,Paragraph,PageBreak,PageTemplate,Spacer,FrameBreak,NextPageTemplate
from reportlab.platypus import BaseDocTemplate, Frame, Paragraph,PageBreak, PageTemplate
from reportlab.platypus import LongTable, TableStyle, BaseDocTemplate, Frame, PageTemplate
from datetime import date, datetime
from textwrap import wrap
from SOA_app.models import *
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MEDIA_DIR = os.path.join(BASE_DIR,'media')
class MyPrint:
    def __init__(self, buffer, pagesize):
        self.buffer = buffer
        if pagesize == 'A4':
            self.pagesize = A4
        elif pagesize == 'Letter':
            self.pagesize = letter
            self.width, self.height = self.pagesize

    @staticmethod
    def print_users(so_num):
        def foot1(canvas,doc):
            width,height = A4
            canvas.saveState()
            canvas.setFont('Times-Roman',9)
            canvas.drawString(width-0.1*inch, 0.1 * inch, "%d" % doc.page)
            canvas.restoreState()
        def foot2(canvas,doc):
            width,height = A4
            canvas.saveState()
            canvas.setFont('Times-Roman',9)
            canvas.drawString(width-0.1*inch, 0.1 * inch,"%d" % doc.page)
        buffer = io.BytesIO()
        doc = BaseDocTemplate(buffer, rightMargin=inch/4,
                                                    leftMargin=inch/4,
                                                    topMargin=inch/2,
                                                    bottomMargin=inch/4,
                                                    pagesize=A4)

        width,height = A4
        contents =[]
        styleSheet = getSampleStyleSheet()
        leftlogoframe = Frame(0.7*cm, height-3.5*cm, 3*cm, 3*cm, id='col1',leftPadding=0, topPadding=0, rightPadding=0, bottomPadding=0,showBoundary = 0)
        TopCenter = Frame(4.7*cm, height-4.5*cm, 5*cm, 4*cm, id='F2',leftPadding=5, topPadding=5, rightPadding=0, bottomPadding=5,showBoundary = 0)
        frameBill = Frame(0.7*cm, height-9*cm, 5*cm, 4.5*cm, id='col1',leftPadding=0, topPadding=0, rightPadding=0, bottomPadding=0,showBoundary = 0)
        frameShip = Frame(6.8*cm, height-9*cm, 5*cm, 4.5*cm, id='col12',leftPadding=0.5, topPadding=0, rightPadding=0, bottomPadding=0,showBoundary = 0)
        frameSales = Frame(12.2*cm, height-7.9*cm, 8.2*cm, 7*cm, id='col12',leftPadding=0, topPadding=0, rightPadding=0, bottomPadding=0,showBoundary = 1)
        frameTable = Frame(0.7*cm, height-25*cm ,19.7*cm, 16* cm,leftPadding=0,bottomPadding=0,rightPadding=0,topPadding=0,showBoundary=1,id='CatBox_frame')
        frame1later = Frame(0.7*cm, height-28*cm ,19.8*cm, 1.7* cm,leftPadding=0,bottomPadding=0,rightPadding=0,topPadding=0,showBoundary = 0,id='col1later')
        framelstpara = Frame(0.7*cm, height-30*cm ,19.5*cm, 1.5* cm,leftPadding=0,bottomPadding=0,rightPadding=0,topPadding=0,showBoundary = 0,id='col1later')
        firstpage = PageTemplate(id='firstpage',frames=[leftlogoframe,TopCenter,frameBill,frameShip,frameSales,frameTable,frame1later,framelstpara],onPage=foot1)
        bodyStyle = ParagraphStyle('Body',fontSize=8)
        contents.append(NextPageTemplate('firstpage'))
        imglogo = MEDIA_DIR + '\logo.jpg'
        logoleft = Image(imglogo)
        logoleft._restrictSize(1.4*inch, 1*inch)
        logoleft.hAlign = 'LEFT'
        logoleft.vAlign = 'CENTER'
        contents.append(logoleft)
        contents.append(FrameBreak())
        ptext1 = '''Brooks Automation, Inc. '''
        ptext12 = '''15 Elizabeth Drive '''
        ptext13 = '''Chelmsford,MA 01824-4111'''
        ptext14 = '''United States'''
        ptext15 = '''Phone: (978) 262-2400'''
        ptext16 =''' Fax:     (978) 262-2500'''
        Uniqvar = sale_order_summery_v.objects.filter(SO_Number=so_num).first()
        OU_org_name = Uniqvar.OU_Organization_Name
        if OU_org_name==None:
            OU_org_name=""
        contents.append(Paragraph(OU_org_name,bodyStyle))
        aladd_c=Uniqvar.Address_o
        if aladd_c==None:
            aladd_c=""
        # al1cty_c=Uniqvar.City_o
        # if al1cty_c==None:
        #     al1cty_c=""
        alzip_c=Uniqvar.ZIP_o
        if alzip_c==None:
            alzip_c=""
        # aladd_c=Uniqvar.Address_c
        # if aladd_c==None:
        #     aladd_c=""
        # aladd_c=Uniqvar.Address_c
        # if aladd_c==None:
        #     aladd_c=""
        # aladd_c=Uniqvar.Address_c
        # if aladd_c==None:
        #     aladd_c=""
        al1_c=Uniqvar.Address_Line1_o
        if al1_c==None:
            al1_c=""
        al2_c=Uniqvar.Address_Line2_o
        if al2_c==None:
            al2_c=""
        al3_c=Uniqvar.Address_Line3_o
        if al3_c==None:
            al3_c=""
        contents.append(Paragraph(aladd_c,bodyStyle))
        contents.append(Paragraph(al1_c,bodyStyle))
        contents.append(Paragraph(al2_c,bodyStyle))
        contents.append(Paragraph(al3_c,bodyStyle))
        # contents.append(Paragraph(al1cty_c,bodyStyle))
        contents.append(Paragraph(alzip_c,bodyStyle))
        # contents.append(Paragraph(al3_c,bodyStyle))
        # contents.append(Paragraph(al3_c,bodyStyle))
        contents.append(FrameBreak())
        ptext31 = '''<b>Bill TO:</b>'''
        # text = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book."
        # wraped_text = "\n".join(wrap(text, 80)) # 80 is line width
        # contents.append(wraped_text)
        contents.append(Paragraph(ptext31,bodyStyle))
        Inv_org_name=Uniqvar.INV_Organization_Name
        if Inv_org_name==None:
            Inv_org_name=""
        contents.append(Paragraph(Inv_org_name,bodyStyle))
        al1add_b=Uniqvar.Address_b
        if al1add_b==None:
            al1add_b=""
        al1_b=Uniqvar.Address_Line1_b
        if al1_b==None:
            al1_b=""
        al2_b=Uniqvar.Address_Line2_b
        if al2_b==None:
            al2_b=""
        al3_b=Uniqvar.Address_Line3_b
        if al3_b==None:
            al3_b=""
        alcity_b=Uniqvar.City_b
        if alcity_b==None:
            alcity_b=""
        aldist_b=Uniqvar.District_s
        if aldist_b==None:
            aldist_b=""
        alstat_b=Uniqvar.State_s
        if alstat_b==None:
            alstat_b=""
        alcnt_b=Uniqvar.County_b
        if alcnt_b==None:
            alcnt_b=""
        al3cont_b=Uniqvar.Country_b
        if al3cont_b==None:
            al3cont_b=""
        al3zip_b=Uniqvar.ZIP_b
        if al3zip_b==None:
            al3zip_b=""
        contents.append(Paragraph(al1add_b,bodyStyle))
        contents.append(Paragraph(al1_b,bodyStyle))
        contents.append(Paragraph(al2_b,bodyStyle))
        contents.append(Paragraph(al3_b,bodyStyle))
        contents.append(Paragraph(alcity_b,bodyStyle))
        contents.append(Paragraph(aldist_b,bodyStyle))
        contents.append(Paragraph(alstat_b,bodyStyle))
        contents.append(Paragraph(alcnt_b,bodyStyle))
        contents.append(Paragraph(al3cont_b,bodyStyle))
        contents.append(Paragraph(al3zip_b,bodyStyle))
        contents.append(FrameBreak())
        ptext41 = '''<b>Ship TO:</b>'''
        contents.append(Paragraph(ptext41,bodyStyle))
        contents.append(Paragraph(Inv_org_name,bodyStyle))
        al1add_s=Uniqvar.Address_s
        if al1add_s==None:
            al1add_s=""
        al1_s=Uniqvar.Address_Line1_s
        if al1_s==None:
            al1_s=""
        al2_s=Uniqvar.Address_Line2_s
        if al2_s==None:
            al2_s=""
        al3_s=Uniqvar.Address_Line3_s
        if al3_s==None:
            al3_s=""
        alcity_s=Uniqvar.City_s
        if alcity_s==None:
            alcity_s=""
        aldist_s=Uniqvar.District_s
        if aldist_s==None:
            aldist_s=""
        alstat_s=Uniqvar.State_s
        if alstat_s==None:
            alstat_s=""
        alcnt_s=Uniqvar.County_s
        if alcnt_s==None:
            alcnt_s=""
        al3cont_s=Uniqvar.Country_s
        if al3cont_s==None:
            al3cont_s=""
        al3zip_s=Uniqvar.ZIP_s
        if al3zip_s==None:
            al3zip_s=""
        contents.append(Paragraph(al1add_s,bodyStyle))
        contents.append(Paragraph(al1_s,bodyStyle))
        contents.append(Paragraph(al2_s,bodyStyle))
        contents.append(Paragraph(al3_s,bodyStyle))
        contents.append(Paragraph(alcity_s,bodyStyle))
        contents.append(Paragraph(aldist_s,bodyStyle))
        contents.append(Paragraph(alstat_s,bodyStyle))
        contents.append(Paragraph(alcnt_s,bodyStyle))
        contents.append(Paragraph(al3cont_s,bodyStyle))
        contents.append(Paragraph(al3zip_s,bodyStyle))
        contents.append(FrameBreak())
        style = getSampleStyleSheet()
        textobject = style["Normal"]
        normal1 = style["Normal"]
        normal1.fontName = "Helvetica-Bold"
        normal1.fontSize = 15
        normal1.spaceBefore = 20
        normal1.spaceAfter = 5
        normal1.leftIndent = 0
        normal1.alignment = TA_CENTER
        normal1.leading = 15
        ptext51 ='''SALES ORDER ACKNOWLEDGEMENT'''
        contents.append(Paragraph(ptext51,normal1))
        al1_sno=Uniqvar.SO_Number
        if al1_sno==None:
            al1_sno=""
        al2_dat=Uniqvar.SO_DATE
        if al2_dat==None:
            al2_dat=""
        al3_cpo=Uniqvar.Customer_PO
        if al3_cpo==None:
            al3_cpo=""
        al4_cur=Uniqvar.Currency
        if al4_cur==None:
            al4_cur=""
        al6_sp=Uniqvar.Employee_FirstName_sp
        if al6_sp==None:
            al6_sp=""
        al6_ssp=Uniqvar.Employee_LastName_sp
        if al6_ssp==None:
            al6_ssp=""
        al8_pt=Uniqvar.Payment_Term
        if al8_pt==None:
            al8_pt=""
        table_story1 = []
        now = al2_dat # current date and time
        year = now.strftime("%Y-%m-%d")
        date_time = now.strftime("%d/%m/%Y")
        # print("date and time:",date_time)
        data1 = [['Sales Order', al1_sno],['Original Order \nDate',date_time],['Customer PO\n',al3_cpo],['Currency\n',al4_cur],['Page', '1 of 1'],['Account Manager',al6_sp,al6_ssp], ['Payment Terms\n Net 30 Days Past Invoice Date\n',al8_pt]]
        t1 = Table(data1,hAlign='CENTER')
        t1.setStyle(TableStyle([
                            # ('FONTSIZE', (0, 0), (-1, 0),22,'Times-Bold'),
                            # ('FONTSIZE', (0, 0), (-1, 0),'Times-Bold'),
                                # ('BOX', (0, 0), (-1, -1), 0.5, colors.black)
                            ('FONTSIZE', (0,0), (-1, -1), 11),
                            ('LINEBELOW',(0,0), (-1,-1),0.25, colors.darkgray),
                            ('INNERGRID',(0,0),(1,7),0.25, colors.black),
                            ('LINEABOVE', (0,0), (-1,-1), 0.25, colors.black),
                            # ('LINEABOVE', (1,0),(0,3), 0.25, colors.black),
                            ('VALIGN',(1,1),(1,-1),'TOP'),
                        ]))
                            # ('LINEABOVE', (0,0), (2,5), 0.25, colors.black),('BOX', (0,0), (-1,-1), 0.15, colors.black),
                            # ('LINEBELOW', (0,0), (-1,-1), 0.25, colors.black),
                            # ('INNERGRID',(1,0),(0,5),0.25, colors.black),

                            #     ('ALIGN',(0,0),(-1,-1),'LEFT'),('HALIGN',(0,0),(-1,-1),'TOP'),
                            #     ('VALIGN',(0,0),(-1,-1),'TOP')]))
        table_story1.append(t1)
        t_keep1 = KeepInFrame(0, 0, table_story1, fakeWidth=False)
        contents.append(t_keep1)
        contents.append(FrameBreak())
        alcc1=Uniqvar.Address_b
        if alcc1==None:
            alcc1=""
        alccphn=Uniqvar.Contact_No_sp
        if alccphn==None:
            alccphn=""
        alccmail=Uniqvar.Email_sp
        if alccmail==None:
            alccmail=""
        al1ccc_name=Uniqvar.Employee_FirstName_csr
        if al1ccc_name==None:
            al1ccc_name=""
        al1ccc_phon=Uniqvar.Contact_No_csr
        if al1ccc_phon==None:
            al1ccc_phon=""
        al1ccc_mail=Uniqvar.Email_csr
        if al1ccc_mail==None:
            al1ccc_mail=""
        table_story21 = []
        data21 = [['Customer Contact','','','','','Customer Care Contact'],[alcc1,'','','','',al1ccc_name],[alccphn,'','','','',al1ccc_phon],[alccmail,'','','','',al1ccc_mail]]
        t21 = LongTable(data21)
        t21.setStyle(TableStyle([
                                        ('ALIGN',(0,1),(-1,-1),'CENTER'),
                                        # ('LINEABOVE',(0,1),(8,1), 0.25, colors.black),
                                        # ('LINEBELOW',(0,1),(8,1), 0.25, colors.darkgray),
                                        # ('INNERGRID',(0,0),(1,0),0.25, colors.black),
                                        # ('BACKGROUND',(0,1),(8,1),'#0F52BA'),
                                        # ('TEXTCOLOR', (0,1), (8, 1), colors.white),
                                        ('FONTSIZE', (0,0), (-1, -1), 11),
                                            ('HALIGN',(0,0),(-1,-1),'TOP'),
                                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                                            # ('GRID',(0,0),(1,0),0.01*inch,(0,0,0,)),
                                        # ('BOX', (0,0), (-1,-1), 0.15, colors.black)
                                        ]))
        table_story21.append(t21)
        t_keep21 = KeepInFrame(0, 0, table_story21, mode='shrink', hAlign='CENTER', vAlign='MIDDLE', fakeWidth=False)
        contents.append(t_keep21)
        table_story2 = []
        Uniqvar = sale_order_summery_v.objects.filter(SO_Number=so_num)
        Transaction_Amount = 0
        Charges = 0
        TAX = 0
        TOTAL =0
        table_story22 =[]
        data22 = [['Line','Item Number/ \nItem description','Customer\nRequest Date','Scheduled Ship\nDate','Qty','UM','Unit Price','Extended\n Price ']]
        t22 = Table(data22,colWidths=(40,125,75,85,40,50,72,72),
                rowHeights=None,)
        t22.setStyle(TableStyle([
                                                ('ALIGN',(0,1),(-1,-1),'CENTER'),
                                                ('LINEABOVE',(0,0),(7,0), 0.25, colors.black),
                                                ('LINEBELOW',(0,0),(7,0), 0.25, colors.darkgray),
                                                ('INNERGRID',(0,0),(7,0),0.25, colors.black),
                                                ('BACKGROUND',(0,0),(7,0),'#0F52BA'),
                                                ('TEXTCOLOR', (0,0), (7, 0), colors.white),
                                                ('FONTSIZE', (0,0), (-1, -1), 11),
                                                    ('HALIGN',(0,0),(-1,-1),'TOP'),
                                                    ('VALIGN',(0,0),(-1,-1),'TOP'),
                                                    ('GRID',(0,1),(-1,-1),0.01*inch,(0,0,0,)),
                                                ('BOX', (0,0), (-1,-1), 0.15, colors.black)
                                                ]))
        table_story22.append(t22)
        t_keep22 = KeepInFrame(0, 0, table_story22, fakeWidth=False)
        # contents.append(t_keep22)
        inx=1
        for Uniqvar_c in Uniqvar:
            al4_qnt1=Uniqvar_c.Quantity
            if al4_qnt1==None:
                    al4_qnt1=""
            al6_upric1=Uniqvar_c.Unit_Price
            if al6_upric1==None:
                    al6_upric1=""
            al1_lin=Uniqvar_c.SO_Number
            if al1_lin==None:
                al1_lin=""
            alno_lin=Uniqvar_c.SO_Line_Number
            if alno_lin==None:
                alno_lin=""
            al11_lin=Uniqvar_c.SO_Line_Shipment
            if al11_lin==None:
                al11_lin=""
            Line1 = alno_lin+'.'+al11_lin
            if Line1==None:
                Line1=""
            al11_ship=Uniqvar_c.Request_DATE
            if al11_ship==None:
                al11_ship=""
            al2_Ino=Uniqvar_c.Item_Number
            if al2_Ino==None:
                al2_Ino=""
            al3_sdat=Uniqvar_c.Schedule_Ship_Date
            if al3_sdat==None:
                al3_sdat=""
            al5_trx=Uniqvar_c.TRX_UOM_Code
            if al5_trx==None:
                al5_trx=""
            al6_upric=Uniqvar_c.Unit_Price
            if al6_upric==None:
                al6_upric=""
            Transaction_Amount = Transaction_Amount + Uniqvar_c.Transaction_Amount
            Charges = Charges + Uniqvar_c.Freight_Charges+Uniqvar_c.Insurance_Charges+Uniqvar_c.CnF_Charges
            TAX = float(TAX) + float(Uniqvar_c.Tax_Amount1)+float(Uniqvar_c.Tax_Amount2)+float(Uniqvar_c.Tax_Amount3)
            TOTAL = float(TOTAL) + float(Uniqvar_c.Transaction_Amount)+float(Uniqvar_c.Freight_Charges+Uniqvar_c.Insurance_Charges+Uniqvar_c.CnF_Charges)+(float(Uniqvar_c.Tax_Amount1)+float(Uniqvar_c.Tax_Amount2)+float(Uniqvar_c.Tax_Amount3))
            nowship = al11_ship # current date and time
            year = nowship.strftime("%Y-%m-%d")
            date_timeship = nowship.strftime("%d/%m/%Y")
            now1 = al3_sdat # current date and time
            year = now1.strftime("%Y-%m-%d")
            date_time1 = now1.strftime("%d/%m/%Y")
            if(inx==1):
                data2 = [['Line','Item Number/ \nItem description','Customer Request\nDate','Scheduled Ship\nDate','Qty','UM','Unit Price','Extended\n Price '],[Line1,al2_Ino,date_time1,date_timeship,al4_qnt1,al5_trx,al6_upric1,al4_qnt1*al6_upric]]
            else:
                data2 = [[Line1,al2_Ino,date_time1,date_timeship,al4_qnt1,al5_trx,al6_upric1,al4_qnt1*al6_upric]]


            t2 = Table(data2,colWidths=(40,95,100,85,40,50,72,72),rowHeights=None,)
            t2.setStyle(TableStyle([('ALIGN',(0,1),(-1,-1),'CENTER'),
                ('LINEABOVE',(0,0),(7,0), 0.25, colors.black),
                ('LINEBELOW',(0,0),(7,0), 0.25, colors.darkgray),
                ('INNERGRID',(0,0),(7,0),0.25, colors.black),
                ('BACKGROUND',(0,0),(7,0),'#0F52BA'),
                ('TEXTCOLOR', (0,0), (7, 0), colors.white),
                ('FONTSIZE', (0,0), (-1, -1), 11),
                    ('HALIGN',(0,0),(-1,-1),'TOP'),
                    ('VALIGN',(0,0),(-1,-1),'TOP'),
                    ('GRID',(0,1),(-1,-1),0.01*inch,(0,0,0,)),
                ('BOX', (0,0), (-1,-1), 0.15, colors.black)
                ]))
            # if float(Line1) <= 15.1:
            #     contents.append(PageBreak())
            # if float(Line1) >=15.5:
            #     contents.append(PageBreak())
            #     float(Line1) = 15.1

            table_story2.append(t2)
            t_keep2 = KeepInFrame(0, 0, table_story2, mode='shrink', hAlign='CENTER', vAlign='MIDDLE', fakeWidth=False)
            inx=inx+1
            if inx==15:
                break
        contents.append(t_keep2)
        contents.append(FrameBreak())
        table_story3 = []
        data3 = [['This document includes an estimate of charges and tax which may differ from the invoice'],['','SUB TOTAL','CHARGES','TAX','TOTAL(USD)'],['',Transaction_Amount,Charges,TAX,TOTAL]]
        t3 = Table(data3)
        t3.setStyle(TableStyle([
                                    ('LINEABOVE', (0,0), (-1,0), 0.25, colors.darkgray),
                                    ('ALIGN',(0,1),(-1,-1),'CENTER'),
                                    ('FONTSIZE', (0,0), (-1, -1), 11),
                                    ('LINEBELOW', (0,0), (-1,0), 0.25, colors.darkgray),
                                    ('LINEABOVE', (0,-1), (-1,-1), 0.25, colors.darkgray),
                                    ('INNERGRID',(1,0),(0,4),0.25, colors.darkgray),
                                    ('LINEBELOW', (0,-1), (-1,-1), 0.25, colors.darkgray),('BOX', (0,0), (-1,-1), 0.15, colors.black) ]))

        table_story3.append(t3)
        t_keep3 = KeepInFrame(0, 0, table_story3, mode='shrink', hAlign='CENTER', vAlign='MIDDLE', fakeWidth=False)
        contents.append(t_keep3)
        contents.append(FrameBreak())
        ptext61 = '''Terms and Conditions can be found at http://www.brooks.com/about/terms-and-conditions'''
        contents.append(Paragraph(ptext61,bodyStyle))
        doc.addPageTemplates([firstpage])
        doc.build(contents)
        pdf = buffer.getvalue()
        buffer.close()
        return pdf
