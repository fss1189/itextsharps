using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using iTextSharp.text;
using iTextSharp.text.pdf;
using IBM.Data.DB2.iSeries;
using L1_PrintWithdrawalLetters.Models;
using Dapper;
using System.Diagnostics;
using System.Threading.Tasks;


namespace readExcel
{  
    
    public partial class Form1 : Form
    {
        string _rootPath = "";

        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);         //program startup path
            System.IO.DirectoryInfo rootPath = System.IO.Directory.GetParent(appPath);  //parent folder of exe program location 
            _rootPath = rootPath.FullName;
        }


        private void confirmButton_Click(object sender, EventArgs e)
        {
            //1.
            Get_SourcePDF_TemplateNameFromAS400(comboBox1.Text);  //form different groups PDF template files (input)

            //2.
            PrintLetterWithGroupAndCountry();
        }


        //1.1
        //read AS400 zvincent.FPL1LTR to form source PDF template     srcPDF= _(SubGroup A1, A2, B1, C, D, E...)_(Country-Code)
        public void Get_SourcePDF_TemplateNameFromAS400(string subgroup)
        {
            List<PDF_FileName> result = new List<PDF_FileName>();  //To form input template PDF filename from output filename (<PDF_FileName>)

            //L1_Letters_chn_template.pdf ==> inCHN_A1提款信_MO_1-18480_page18480_1.pdf (from A1提款信_MO_1-18480_page18480_1.pdf)
            //L1_Letters_por_template.pdf ==> inPOR_A1提款信_MO_1-18480_page18480_1.pdf (from A1提款信_MO_1-18480_page18480_1.pdf)
            //PDF多於5000頁, 分5000頁為一組, 以p1,p2,p3...開頭 (eg. p1_inCHN_A1提款信_MO_1-18480_page18480_1.pdf, p2_inCHN_A1提款信_MO_1-18480_page18480_1.pdf...p1_inPOR_A1提款信_MO_1-18480_page18480_1.pdf...) 
            result = Get_PDF_FilenameFromAS400(comboBox1.Text).ToList();  //get output PDF filename (for form input template PDF)

            if (result.Count() > 0)
            {
                foreach (var itm in result)  //生成不同的source template pdf files (for multi-threads using)
                {
                    // Will not overwrite if the destination file already exists.  A1_MO, A1_TP....., E_MO, E_TP...
                    if (itm.LTRSUBGRP.Trim().ToUpper() == "E")   //for E group only
                        File.Copy(GlobalVar.wPATH() + GlobalVar.src_templatePDF_E(), GlobalVar.wPATH() + "input_" + itm.FILENAME);    //E_Letters_template.pdf ==> input_E提款信_MO_1-120_page120_1.pdf)
                    else
                    {   //for A, B, C, D group only
                        File.Copy(GlobalVar.wPATH() + GlobalVar.src_templatePDF_L1_chn(), GlobalVar.wPATH() + "inCHN_" + itm.FILENAME); //L1_Letters_chn_template.pdf ==> inCHN_A1提款信_MO_1-18480_page18480_1.pdf (from A1提款信_MO_1-18480_page18480_1.pdf)
                        File.Copy(GlobalVar.wPATH() + GlobalVar.src_templatePDF_L1_por(), GlobalVar.wPATH() + "inPOR_" + itm.FILENAME); //L1_Letters_por_template.pdf ==> inPOR_A1提款信_MO_1-18480_page18480_1.pdf (from A1提款信_MO_1-18480_page18480_1.pdf)
                    }
                }
            }
        }

        //2.
        #region [PDF letters]
        //Form different-PDF-File with different subgroup and country except group E letters ___________________________ All
        public void PrintLetterWithGroupAndCountry()
        {
            List<Task> TLists = new List<Task>();
       
            PcListBox.Items.Add("First Starting in " + DateTime.Now.ToString() + " for adding to GlobalVar.FPL1LTR_List");
            PcListBox.Items.Add("");

            GlobalVar.FPL1LTR_List = Get_GroupABCDE_dataFromAS400(comboBox1.Text);  //sub-set data (subGroup A1,A2,B1,B2,C1,C2,D,E data)

            Task tasks;
            foreach (var fItem in Get_PDF_FilenameFromAS400(comboBox1.Text).ToList())  //get diff subgroup and country to form PDF file name (For output)
            {
                //if (fItem.FILENAME.Contains("1-18474")) //..........................
                {
                    if (fItem.LTRSUBGRP.Trim() == "E")
                    {
                        //PrintGroupWithCountry_E(fItem.FILENAME.Trim(), fItem.LTRSUBGRP.Trim(), fItem.PTNCNY.Trim());  //E Letter is bi-lang PDF
                        tasks = Task.Factory.StartNew(() => PrintGroupWithCountry_E(fItem.FILENAME.Trim(), fItem.LTRSUBGRP.Trim(), fItem.PTNCNY.Trim()));  //E Letter is bi-lang PDF
                        TLists.Add(tasks);
                    }
                    else
                    {
                        //PrintGroupWithCountry("", fItem.FILENAME.Trim(), fItem.LTRSUBGRP.Trim(), fItem.PTNCNY.Trim(), fItem.PAGE);
                        tasks = Task.Run(() => PrintGroupWithCountry("", fItem.FILENAME.Trim(), fItem.LTRSUBGRP.Trim(), fItem.PTNCNY.Trim(), fItem.PAGE));
                        TLists.Add(tasks);
                    }
                }
           }

            //Wait for all of the tasks to complete (WaitAny.ToArray()任一TASK完成,程式結束,其他Task仍在後台運作)
            Task.WaitAll(TLists.ToArray());

            PcListBox.Items.Add("Finished in " + DateTime.Now.ToString());
        }


        //2.1
        //Form different-PDF-File with different subgroup and country except group E letters ___________________________ 2.1
        private void PrintGroupWithCountry(string inputHead, string outputPDFFile, string LtrSubGrp, string PtnCNY, string pageFlag)   //pageFlag is single or double page (PDF output) : A,B,C,D letters
        {
            //按AS400系統的短訊語言分別生成中文或葡文信函 (SMS Lang. C=中文信, P=葡文信)
            //- 如沒有語言選項，且只有葡文姓名，同時生成中文及葡文信 (2 letters 中及葡文信)
            //- 承上點，即如果沒有語言選項但有中文姓名，生成中文信 (1 letter 中文信)

            var ABCD_Group_List = GlobalVar.FPL1LTR_List.Where(grp => grp.LTRSUBGRP.Trim() == LtrSubGrp && grp.PTNCNY.Trim() == PtnCNY && grp.PAGE == pageFlag).ToList();

            int idx=0;
            //has data
            if (ABCD_Group_List.Count() > 0)
            {
                using (iTextSharp.text.Document doc = new iTextSharp.text.Document(PageSize.A4))
                {
                    string prefix_breakPage = "";
                    prefix_breakPage = ABCD_Group_List.Count() > 5000 ? textBoxFrom.Text + "_" : "";
                    // Create a PdfStreamWriter instance, responsible to write the document into the specified file
                    using (var fs = new FileStream(GlobalVar.wPATH() + prefix_breakPage + outputPDFFile, FileMode.Create)) //output result PDF file
                    {
                        using (var copy = new PdfSmartCopy(doc, fs))
                        {
                            doc.Open();

                            foreach (var pageInfo in ABCD_Group_List)  //passing Subgroup and Country-code for reading data (A or B or C or D group list with diff. country) 
                            {
                                idx++;
                                if (pageInfo.PAGE == "2")  //double-page
                                {
                                    FillTextToPDF_CHN("inCHN_" + outputPDFFile, copy, pageInfo, LtrSubGrp, PtnCNY); //chinese version (rename output filename to input file name)  inCHN_xxx
                                    FillTextToPDF_POR("inPOR_" + outputPDFFile, copy, pageInfo, LtrSubGrp, PtnCNY); //portuguese version (rename output filename to input file name) inPOR_xxx
                                }
                                else
                                {
                                    if (pageInfo.PTSLNG.Trim() == "C")  //Chinese
                                        FillTextToPDF_CHN("inCHN_" + outputPDFFile, copy, pageInfo, LtrSubGrp, PtnCNY);   //chinese version
                                    else
                                        if (pageInfo.PTSLNG.Trim() == "P") //Portuguese
                                            FillTextToPDF_POR("inPOR_" + outputPDFFile, copy, pageInfo, LtrSubGrp, PtnCNY);   //portuguese version
                                        else
                                            if (pageInfo.PTNCNA.Trim() != "")  //沒有PTSLNG, 如有中文名 //Chinese letter
                                                FillTextToPDF_CHN("inCHN_" + outputPDFFile, copy, pageInfo, LtrSubGrp, PtnCNY);   //chinese version
                                }

                                if (idx % 500 == 0)
                                    Console.WriteLine(idx.ToString() + "  " + pageInfo.LTRNO + " === " + DateTime.Now.ToString());
                            }
                        } //end of copy
                    }
                }
            }
        }
        
        //2.2.1
        //fill text to each PDF except group E letters__________________________________________________________________ 2.2.1
        private void FillTextToPDF_CHN(string inputPDFFile, PdfSmartCopy copy, FPL1LTR pageInfo, string LtrSubGrp, string PtnCNY)
        {
           // Console.WriteLine("start chn  " + pageInfo.LTRNO + " === " + DateTime.Now.ToString());

            using (var baos = new MemoryStream())  //p1_inCHN_xxxxxx.PDF
            {
                //using (PdfReader templateReader = new PdfReader(outputPDFFile))  //source template PDF (chn version)  d:\L1Letters\L1_Letters_chn_template.pdf 
                using (PdfReader templateReader = new PdfReader(GlobalVar.wPATH() + inputPDFFile.Replace("inXXX_","inCHN_")))  //input= "d:\L1Letters\" + "p1_inCHN_A1提款信_MO_1-18480_page18480_1.PDF"
                {
                    using (PdfStamper stamper = new PdfStamper(templateReader, baos))   //convert inXXX_ to inCHN_
                    {
                        //create a font family (細明體HKSCS + EUDC)
                        string fontPath = @"d:\MINGLIU.TTC,2";   //path + @"\MINGLIU.TTC,2";
                        BaseFont bfChinese = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        string fontPath2 = @"d:\EUDC.TTF";
                        BaseFont bfChinese2 = BaseFont.CreateFont(fontPath2, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                        List<BaseFont> FontFamily = new List<BaseFont>();
                        FontFamily.Add(bfChinese);
                        FontFamily.Add(bfChinese2);

                        AcroFields fields = stamper.AcroFields;
                        fields.SubstitutionFonts = FontFamily; 
                        fields.GenerateAppearances = true;

                        fields.SetField("cname", pageInfo.PTNCNA);
                        stamper.AcroFields.SetFieldProperty("pname", "textsize", 0f, null);  //font auto resize (0f)
                        fields.SetField("pname", pageInfo.PTNNAM);
                        fields.SetField("ltrno", pageInfo.LTRNO);
                        fields.SetField("idno", pageInfo.PTNIDN.Substring(0, 4) + "XXXX");
                        fields.SetField("ofcdate", DateTime.Now.ToString("dd/MM/yyyy"));
                        fields.SetField("ptnno", pageInfo.PTNNO);
                        fields.SetField("idx", pageInfo.IDX.ToString());
                        fields.SetField("ltrgrp", pageInfo.LTRGRP);

                        stamper.AcroFields.SetFieldProperty("addr1", "textsize", 0f, null);
                        fields.SetField("addr1", pageInfo.ADDR1);

                        if (pageInfo.ADDR2.Trim() == "")   //向前移 (如空白, 分段地址向上移)
                            if (pageInfo.ADDR3.Trim() == "")
                            {
                               pageInfo.ADDR2 = pageInfo.ADDR4 ;
                               pageInfo.ADDR4 = "" ;
                            }
                            else
                            {
                               pageInfo.ADDR2 = pageInfo.ADDR3 ;
                               pageInfo.ADDR3 = pageInfo.ADDR4 ;
                               pageInfo.ADDR4 = "" ;
                            }

                        if (pageInfo.ADDR3.Trim() == "")  //向前移 (如空白, 分段地址向上移)
                        {
                            pageInfo.ADDR3 = pageInfo.ADDR4 ;
                            pageInfo.ADDR4 = "" ;
                        }

                        stamper.AcroFields.SetFieldProperty("addr2", "textsize", 0f, null);
                        fields.SetField("addr2", pageInfo.ADDR2.Trim());
                        stamper.AcroFields.SetFieldProperty("addr3", "textsize", 0f, null);
                        fields.SetField("addr3", pageInfo.ADDR3.Trim());
                        fields.SetField("addr4", pageInfo.ADDR4);
                        

                        if (pageInfo.LTRGRP.Trim() == "D")  //3.3 政府管理子帳戶結餘, 3.4 提取款項原因
                        {
                            fields.SetField("p1balance", "可提取的款項為澳門幣" + String.Format("{0:#,#00.00}", pageInfo.TXNBAL) + "（已包括" + DateTime.Now.Year.ToString() + "年度的款項）");
                            stamper.AcroFields.SetFieldProperty("p2age", "textsize", 0f, null);  //font auto resize (0f)
                            fields.SetField("p2age", "未滿65歲，正收取社會保障基金殘疾金超過一年或社會工作局特別殘疾津貼");
                        }
                        else
                        {
                            fields.SetField("p1balance", "結餘為澳門幣" + String.Format("{0:#,#00.00}", pageInfo.TXNBAL) + "（已包括" + DateTime.Now.Year.ToString() + "年度的款項）");
                            fields.SetField("p2age", "已年滿65歲");
                        }

                        if (pageInfo.LTRGRP.Trim() == "C")  //3.5 款項發放
                        {
                            fields.SetField("p4bank1", "存入本人澳門幣銀行帳號＿＿＿＿＿＿＿＿＿＿銀行名稱＿＿＿＿＿＿＿");  //全型underline
                            fields.SetField("p4bank2", "（須附同澳門居民身份證影印本及銀行帳號影印本）");
                            fields.SetField("p4bank3", "");
                        }
                        else
                        {
                            fields.SetField("p4bank1", "款項將按以下順序存入申請人的銀行帳戶：1.收取社會保障基金養老金或");
                            fields.SetField("p4bank2", "殘疾金的銀行帳戶、2.收取社會工作局敬老金的銀行帳戶、3.收取社會工");
                            fields.SetField("p4bank3", "作局特別殘疾津貼的銀行帳戶。");
                        }

                        if (pageInfo.LTRGRP.Trim()=="A")
                        {
                            fields.SetField("p6point", "6.");
                            fields.SetField("p6label", "自動提款登記：");

                            fields.SetField("p6auto1", "同意參與《自動提款登記》（詳見自動提款登記指南），並知悉自登");
                            fields.SetField("p6auto2", "記翌年起在符合條件的年度可無需辦理提款申請，相關年度的分配款");
                            fields.SetField("p6auto3", "項將發放至上述的銀行帳戶。");

                            string fontPath3 = @"c:\windows\fonts\Webdings.TTF";
                            BaseFont bfChinese3 = BaseFont.CreateFont(fontPath3, BaseFont.IDENTITY_V, BaseFont.EMBEDDED);
                            List<BaseFont> FontFamily3 = new List<BaseFont>();
                            FontFamily3.Add(bfChinese3);
                            AcroFields fields3 = stamper.AcroFields;
                            fields3.SubstitutionFonts = FontFamily3;
                            fields3.SetField("p6chkbox", "");

                            _GenQRCode(stamper, "QRcode", pageInfo.PTNNO);
                        }

                        stamper.FormFlattening = true;
                    }

                    using (var template_filled = new PdfReader(baos.ToArray()))
                    {
                        copy.AddPage(copy.GetImportedPage(template_filled, 1));
                    }
                }
            } // end : fill text in a PDF 

            //Console.WriteLine("end chn  " + pageInfo.LTRNO + " === " + DateTime.Now.ToString());
        }

        //2.2.1
        //fill text to each PDF except group E letters__________________________________________________________________ 2.2.2
        private void FillTextToPDF_POR(string inputPDFFile, PdfSmartCopy copy, FPL1LTR pageInfo, string LtrSubGrp, string PtnCNY)
        {
            using (var baos = new MemoryStream())  //inputHead=p1_inPOR_xxxxxx.PDF
            {
                using (PdfReader templateReader = new PdfReader(GlobalVar.wPATH() + inputPDFFile.Replace("inXXX_", "inPOR_")))  //input= "d:\L1Letters\" + "inPOR_A1提款信_MO_1-18480_page18480_1.PDF"
                {
                    using (PdfStamper stamper = new PdfStamper(templateReader, baos))
                    {
                        //create a font family (細明體HKSCS + EUDC)
                        string fontPath = @"d:\MINGLIU.TTC,2";   //path + @"\MINGLIU.TTC,2";
                        BaseFont bfChinese = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        string fontPath2 = @"d:\EUDC.TTF";
                        BaseFont bfChinese2 = BaseFont.CreateFont(fontPath2, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                        List<BaseFont> FontFamily = new List<BaseFont>();
                        FontFamily.Add(bfChinese);
                        FontFamily.Add(bfChinese2);

                        AcroFields fields = stamper.AcroFields;
                        fields.SubstitutionFonts = FontFamily;   
                        fields.GenerateAppearances = true;

                        fields.SetField("cname", pageInfo.PTNCNA);
                        stamper.AcroFields.SetFieldProperty("pname", "textsize", 0f, null);  //font auto resize (0f)
                        fields.SetField("pname", pageInfo.PTNNAM);
                        fields.SetField("ltrno", pageInfo.LTRNO);
                        fields.SetField("idno", pageInfo.PTNIDN.Substring(0, 4) + "XXXX");
                        fields.SetField("ofcdate", DateTime.Now.ToString("dd/MM/yyyy"));
                        fields.SetField("ptnno", pageInfo.PTNNO);
                        fields.SetField("idx", pageInfo.IDX.ToString());
                        fields.SetField("ltrgrp", pageInfo.LTRGRP);

                        stamper.AcroFields.SetFieldProperty("addr1", "textsize", 0f, null);
                        fields.SetField("addr1", pageInfo.ADDR1);

                        if (pageInfo.ADDR2.Trim() == "")   //向前移 (如空白, 分段地址向上移)
                            if (pageInfo.ADDR3.Trim() == "")
                            {
                                pageInfo.ADDR2 = pageInfo.ADDR4;
                                pageInfo.ADDR4 = "";
                            }
                            else
                            {
                                pageInfo.ADDR2 = pageInfo.ADDR3;
                                pageInfo.ADDR3 = pageInfo.ADDR4;
                                pageInfo.ADDR4 = "";
                            }

                        if (pageInfo.ADDR3.Trim() == "")  //向前移 (如空白, 分段地址向上移)
                        {
                            pageInfo.ADDR3 = pageInfo.ADDR4;
                            pageInfo.ADDR4 = "";
                        }

                        stamper.AcroFields.SetFieldProperty("addr2", "textsize", 0f, null);
                        fields.SetField("addr2", pageInfo.ADDR2.Trim());
                        stamper.AcroFields.SetFieldProperty("addr3", "textsize", 0f, null);
                        fields.SetField("addr3", pageInfo.ADDR3.Trim());
                        fields.SetField("addr4", pageInfo.ADDR4);


                        if (pageInfo.LTRGRP.Trim() == "D")  //3.3 政府管理子帳戶結餘, 3.4 提取款項原因   (未滿65歲)
                        {
                            if (pageInfo.UPLBAL < pageInfo.TXNBAL)  //未滿65歲 (LA), 提款上限 (UPLBAL) < 結餘 (TXNBAL) then uplbal else txnbal
                                fields.SetField("p1balance", "O montante máximo a levantar é de MOP$" + String.Format("{0:#,#00.00}", pageInfo.UPLBAL) + " (incluindo a verba do ano " + DateTime.Now.Year.ToString() + ").");
                            else
                                fields.SetField("p1balance", "O montante máximo a levantar é de MOP$" + String.Format("{0:#,#00.00}", pageInfo.TXNBAL) + " (incluindo a verba do ano " + DateTime.Now.Year.ToString() + ").");
                            fields.SetField("p2age", "Não ter completado 65 anos de idade, e estar a receber a pensão de invalidez do Fundo de Segurança Social há mais de um ano ou o subsídio de invalidez especial do Instituto de Acção Social.");
                        }
                        else
                        {
                            fields.SetField("p1balance", "O saldo é de MOP$" + String.Format("{0:#,#00.00}", pageInfo.TXNBAL) + " (incluindo a verba do ano " + DateTime.Now.Year.ToString() + ").");
                            fields.SetField("p2age", "Ter completado 65 anos de idade.");
                        }


                        if (pageInfo.LTRGRP.Trim() == "C")  //3.5 款項發放
                        {
                            //fields.SetField("p4bank1", "Por depósito na minha conta bancária em MOP n.º＿＿＿＿＿Banco＿＿＿"); //全型underline
                            fields.SetField("p4bank1", "Por depósito na conta bancária em MOP nº＿＿＿＿＿＿＿＿＿Banco＿＿＿"); //全型underline
                            fields.SetField("p4bank2", "(Deve anexar fotocópias do BIRM e da conta bancária)");
                            fields.SetField("p4bank3", "");
                            fields.SetField("p4bank4", "");
                            fields.SetField("p4bank5", "");
                        }
                        else
                        {
                            fields.SetField("p4bank1", "A  verba  será   depositada   na  conta  bancária   do  requerente   pela  ordem");
                            fields.SetField("p4bank2", "seguinte: 1. Conta bancária na qual recebe a pensão para idosos ou pensão de");
                            fields.SetField("p4bank3", "invalidez do FSS. 2. Conta bancária na qual recebe o subsídio para idosos do");
                            fields.SetField("p4bank4", "Instituto de  Acção  Social.  3. Conta  bancária  na  qual recebe o subsídio de");
                            fields.SetField("p4bank5", "invalidez especial do Instituto de Acção Social.");
                        }

                        if (pageInfo.LTRGRP.Trim() == "A")
                        {
                            fields.SetField("p6point", "6.");
                            fields.SetField("p6label", "Inscrição    de     levantamento");
                            fields.SetField("p6label2", "automático de verbas:");

                            fields.SetField("p6auto1", "Se   concordar ,   pode    participar   na   “Inscrição   de    levantamento");
                            fields.SetField("p6auto2", "automático  de  verbas”  (vide  Guia para a  inscrição de  levantamento");
                            fields.SetField("p6auto3", "automático  de  verbas), tomando conhecimento  de que a partir do ano");
                            fields.SetField("p6auto4", "seguinte   ao   registo,  não   é   preciso   efectuar  o   requerimento   de");
                            fields.SetField("p6auto5", "levantamento  de  verba desde que sejam  preenchidos os requisitos  no");
                            fields.SetField("p6auto6", "ano  em  causa.  A  verba  do  ano  relevante  será  atribuída   na   conta");
                            fields.SetField("p6auto7", "bancária acima referida.");

                            string fontPath3 = @"c:\windows\fonts\Webdings.TTF";
                            BaseFont bfChinese3 = BaseFont.CreateFont(fontPath3, BaseFont.IDENTITY_V, BaseFont.EMBEDDED);
                            List<BaseFont> FontFamily3 = new List<BaseFont>();
                            FontFamily3.Add(bfChinese3);
                            AcroFields fields3 = stamper.AcroFields;
                            fields3.SubstitutionFonts = FontFamily3;
                            fields3.SetField("p6chkbox", "");

                            _GenQRCode(stamper, "QRcode", pageInfo.PTNNO);
                        }

                        stamper.FormFlattening = true;
                    }

                    using (var template_filled = new PdfReader(baos.ToArray()))
                    {
                        copy.AddPage(copy.GetImportedPage(template_filled, 1));
                    }
                }
            } // end : fill text in a PDF 
        }

        //2.3
        //print QRcode in button field
        private void _GenQRCode(PdfStamper stamper, string FieldName, string QrCodeContent)
        {
            AcroFields form = stamper.AcroFields;
            PushbuttonField ad = form.GetNewPushbuttonFromField(FieldName);
            BarcodeQRCode qrcode = new BarcodeQRCode(QrCodeContent, 2, 2, null);
            var qrcodeImg = qrcode.GetImage();
            ad.Layout = PushbuttonField.LAYOUT_ICON_ONLY;
            ad.ProportionalIcon = true;
            ad.Image = qrcodeImg;

            form.ReplacePushbuttonField(FieldName, ad.Field);
        }

        //2.4
        //Form PDF-File with subgroup E and country only _______________________________________________________________ E.1   (E letters 是中葡文template)
        private void PrintGroupWithCountry_E(string outputPDFFile, string LtrSubGrp, string PtnCNY)
        {
            var E_Group_List = GlobalVar.FPL1LTR_List.Where(grp => grp.LTRSUBGRP.Trim() == LtrSubGrp && grp.PTNCNY.Trim() == PtnCNY).ToList();

            //has data
            if (E_Group_List.Count() > 0)
            {
                using (var fs = new FileStream(GlobalVar.wPATH() + outputPDFFile, FileMode.Create))
                {
                    iTextSharp.text.Document doc = new iTextSharp.text.Document(PageSize.A4);
                    PdfSmartCopy copy = new PdfSmartCopy(doc, fs);
                    doc.Open();

                    foreach (var pageInfo in E_Group_List)
                    {
                        FillTextToPDF_E("input_" +　outputPDFFile, copy, pageInfo, LtrSubGrp, PtnCNY);  //group E only
                    }

                    copy.Close();
                    doc.Close();
                }
            }
        }

        //2.4.1
        //fill text to each PDF withgroup E letters_____________________________________________________________________ E.2
        private void FillTextToPDF_E(string inputPDFFile, PdfSmartCopy copy, FPL1LTR pageInfo, string LtrSubGrp, string PtnCNY)
        {
            using (var baos = new MemoryStream())
            {
                using (PdfReader templateReader = new PdfReader(GlobalVar.wPATH() + inputPDFFile))  //source template PDF (group E only)  d:\L1Letters\E_Letters_template.pdf
                {
                    using (PdfStamper stamper = new PdfStamper(templateReader, baos))
                    {
                        //create a font family (細明體HKSCS + EUDC)
                        string fontPath = @"d:\MINGLIU.TTC,2";   //path + @"\MINGLIU.TTC,2";
                        BaseFont bfChinese = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                        string fontPath2 = @"d:\EUDC.TTF";
                        BaseFont bfChinese2 = BaseFont.CreateFont(fontPath2, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                        List<BaseFont> FontFamily = new List<BaseFont>();
                        FontFamily.Add(bfChinese);
                        FontFamily.Add(bfChinese2);

                        AcroFields fields = stamper.AcroFields;
                        fields.SubstitutionFonts = FontFamily;
                        fields.GenerateAppearances = true;

                        fields.SetField("cname", pageInfo.PTNCNA);
                        stamper.AcroFields.SetFieldProperty("pname", "textsize", 0f, null);  //font auto resize (0f)
                        fields.SetField("pname", pageInfo.PTNNAM);
                        fields.SetField("ltrno", pageInfo.LTRNO);
                        fields.SetField("idno", pageInfo.PTNIDN.Substring(0, 4) + "XXXX");
                        fields.SetField("ofcdate", DateTime.Now.ToString("dd/MM/yyyy"));
                        fields.SetField("ptnno", pageInfo.PTNNO);
                        fields.SetField("ptnnoidx", pageInfo.PTNNO.Trim() + "/" + pageInfo.IDX.ToString().Trim());
                        fields.SetField("barcode", pageInfo.PTNNO + "-" + pageInfo.LTRNO);

                        stamper.AcroFields.SetFieldProperty("addr1", "textsize", 0f, null);
                        fields.SetField("addr1", pageInfo.ADDR1);

                        if (pageInfo.ADDR2.Trim() == "")   //向前移 (如空白, 分段地址向上移)
                            if (pageInfo.ADDR3.Trim() == "")
                            {
                                pageInfo.ADDR2 = pageInfo.ADDR4;
                                pageInfo.ADDR4 = "";
                            }
                            else
                            {
                                pageInfo.ADDR2 = pageInfo.ADDR3;
                                pageInfo.ADDR3 = pageInfo.ADDR4;
                                pageInfo.ADDR4 = "";
                            }

                        if (pageInfo.ADDR3.Trim() == "")  //向前移 (如空白, 分段地址向上移)
                        {
                            pageInfo.ADDR3 = pageInfo.ADDR4;
                            pageInfo.ADDR4 = "";
                        }

                        stamper.AcroFields.SetFieldProperty("addr2", "textsize", 0f, null);
                        fields.SetField("addr2", pageInfo.ADDR2.Trim());
                        stamper.AcroFields.SetFieldProperty("addr3", "textsize", 0f, null);
                        fields.SetField("addr3", pageInfo.ADDR3.Trim());
                        fields.SetField("addr4", pageInfo.ADDR4);

                        //_GenQRCode(stamper, "QRcode", pageInfo.PTNNO);     //E-Letters don't need QRcode! Requested by Elyse on 17/7/2019

                        stamper.FormFlattening = true;
                    }

                    using (var template_filled = new PdfReader(baos.ToArray()))
                    {
                        copy.AddPage(copy.GetImportedPage(template_filled, 1));
                    }
                }
            } // end : fill text in a PDF (group E only)
        }
        #endregion


        //read full set data
        public List<FPL1LTR> Get_GroupABCDE_dataFromAS400(string subgroup)  //without parameters
        {
            List<FPL1LTR> result = new List<FPL1LTR>();

            //SELECT PTNNO, PTNCNY, PTNMACAU, LTRNO, LTRGRP, LTRSUBGRP FROM zvincent.fpl1ltr WHERE ltrgrp<>'' ORDER BY LTRGRP,LTRSUBGRP,PTNMACAU,PTNCNY,PTNNO
            string sqlStr = @"select * from zvincent.fpl1ltr WHERE ltrgrp<>'' and ltrsubgrp=@subgroup_P and (idx between @pFrom and @pTo) ORDER BY PTNMACAU,PTNCNY,PAGE,PTNNO for read only ";

            using (iDB2Connection conn = new iDB2Connection(GlobalVar.getAS400ConnString())) //get as400 conn string in as400Model
            {
                conn.Open();
                result = conn.Query<FPL1LTR>(sqlStr, new { subgroup_P = subgroup, pFrom = Convert.ToInt16(textBoxFrom.Text), pTo = Convert.ToInt32(textBoxTo.Text) }).ToList();  //To list (L1 letters)

            } //------- end

            return result;
        }


        //read AS400 zvincent.FPL1LTR to form FILENAME.pdf (Group A, B, C, D, E) output
        public List<PDF_FileName> Get_PDF_FilenameFromAS400(string subgroup)  //output PDF filename
        {
            List<PDF_FileName> result = new List<PDF_FileName>();

            string sqlStr = @"SELECT PAGENO,PAGE,LTRSUBGRP,PTNCNY,trim(LTRSUBGRP)||'提款信'||'_'||ptncny||'_'||trim(frmIdx)||'-'||trim(toIdx)||'_'||'page'||pageNo||'_'||PAGE||'.PDF' FileName 
                              From (SELECT ltrsubgrp,ptnmacau,ptncny,min(idx) frmIdx, max(idx) toidx,max(idx) - min(idx) + 1 pageNo, PAGE 
                                    from (SELECT case when length(trim(ltrsubgrp))=1 then integer(substr(ltrno,10)) else integer(substr(ltrno,11)) end idx,ltrsubgrp,ptnmacau,ptncny,PAGE
                                          FROM zvincent.FPL1LTR             
                                          where ltrgrp <> '' and ltrsubgrp=@subgroup_P) a                                         
                                          group by ltrsubgrp,ptnmacau,ptncny,PAGE) a                       
                              order by ltrsubgrp,ptnmacau,ptncny,PAGE FOR READ ONLY";
            
            using (iDB2Connection conn = new iDB2Connection(GlobalVar.getAS400ConnString())) //get as400 conn string in as400Model
            {
                conn.Open();
                result = conn.Query<PDF_FileName>(sqlStr, new { subgroup_P = subgroup }).ToList();   //(pdf letters)

            } //------- end

            return result;
        }


        

                         

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string[] pdfList = Directory.GetFiles(@"d:\L1Letters\", "*.pdf");  //clear old pdf files
            foreach (string srcPDF in pdfList)
            {
                if (!srcPDF.ToLower().Contains("template"))   //delete all pdf except template pdf
                    File.Delete(srcPDF);
            }
            MessageBox.Show("Finished to delete !","Message");
        }

       

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }  // end of Form
}
