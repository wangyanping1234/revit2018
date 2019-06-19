//
// (C) Copyright 2003-2017 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

using System;
using System.IO;
using System.Collections;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;

using Autodesk;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using System.Drawing;
using System.Globalization;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;

namespace Revit.SDK.Samples.CreateDimensions.CS
{
    /// <summary>
    /// Implements the Revit add-in interface IExternalCommand
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class Command : IExternalCommand
    {
        ExternalCommandData m_revit = null;    //store external command 存外部命令
        string m_errorMessage = " ";           // store error message 存错误信息
        ArrayList m_walls = new ArrayList();   //store the wall of selected 存选择的墙
        const double precision = 0.0000001;   //store the precision 存精度   

        /// <summary>
        /// Implement this method as an external command for Revit.  作为revit 的外部命令 实现此方法
        /// </summary>
        /// <param name="commandData">An object that is passed to the external application 
        /// which contains data related to the command, 
        /// such as the application object and active view.</param>
        /// <param name="message">A message that can be set by the external application 
        /// which will be displayed if a failure or cancellation is returned by 
        /// the external command.</param>
        /// <param name="elements">A set of elements to which the external application 
        /// can add elements that are to be highlighted in case of failure or cancellation.</param>
        /// <returns>Return the status of the external command. 
        /// A result of Succeeded means that the API external method functioned as expected. 
        /// Cancelled can be used to signify that the user cancelled the external operation 
        /// at some point. Failure should be returned if the application is unable to proceed with 
        /// the operation.</returns>
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            try
            {
                m_revit = revit;
                UIDocument uidoc = revit.Application.ActiveUIDocument;
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                string info = "Selected elements:\n";
                foreach (ElementId id in selectedIds)
                {
                    Element elem = uidoc.Document.GetElement(id);
                    info += "type->"+elem.GetType().ToString()+"\n";
                    info += "name->"+elem.Name + "\n";
                    Rebar bar = elem as Rebar;
                    
                    if (bar!= null)
                    {
                        IList<Subelement> elem2 = bar.GetSubelements();
                        //info+="Subelement->counts="+elem2.Count.ToString()+"\n";
                        
                        RebarShape rebshape = uidoc.Document.GetElement(bar.RebarShapeId) as RebarShape;
                        ParameterSet para=bar.Parameters;
                        Parameter dm=bar.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER);

                        if (rebshape != null)
                        {
                            //钢筋形状线，用这个绘制Bitmap
                            Bitmap bmp = new Bitmap(50 * 4 * 2, 50 * 4 * 2);
                            Pen pen = new Pen(System.Drawing.Color.Black, 4);
                            Font myFont = new Font("宋体", 10, FontStyle.Bold);
                            Brush bush = new SolidBrush(System.Drawing.Color.Black);//填充的颜色
                            NumberFormatInfo m_nfi = new NumberFormatInfo();

                            //m_nfi = new CultureInfo("de-DE", false).NumberFormat;
                            m_nfi.NumberDecimalSeparator = ".";
                            m_nfi.NumberDecimalDigits = 0;
                            m_nfi.NumberGroupSeparator = "";
                            Graphics ht = Graphics.FromImage(bmp);
                            PointF pt1 = new PointF(0, 0), pt2 = new PointF(0, 0);
                            PointF pt1_r = new PointF(0, 0), pt2_r = new PointF(0, 0);
                            //获取长度
                            IList<Curve> rlines = bar.GetShapeDrivenAccessor().ComputeDrivingCurves();
                            List<double> linelengths = new List<double>();
                            int i = 0;
                            double reallen = 0.0;
                            foreach (Line line in rlines)
                            {
                                //计算实际长度，常量304.8

                                if ((i == 0) || (i == rlines.Count - 1))
                                {
                                    reallen = line.Length + 0.5*dm.AsDouble();
                                    
                                }
                                else {
                                    reallen = line.Length + dm.AsDouble();
                                }
                                reallen *= 304.8;
                                linelengths.Insert(i, reallen);

                                i++;
                            }
                            IList<Curve> lines = rebshape.GetCurvesForBrowser();
                            i = 0;
                            foreach (Line line in lines)
                            {
                                IList<XYZ> xyzs = line.Tessellate();
                                //info += "\nline.Length=" + line.Length*25.4;
                                //info += "\npt1=" + xyzs[0].ToString();
                                //info += "\npt2=" + xyzs[1].ToString();
                                //移轴，防止出现负的坐标
                                pt1.X = (float)(xyzs[0].X*25.4+200);
                                pt1.Y = (float)(xyzs[0].Y*25.4+200);
                                pt2.X = (float)(xyzs[1].X*25.4+200);
                                pt2.Y = (float)(xyzs[1].Y*25.4+200);
                                double  jiaodu=0;
                                ht.DrawLine(pen, pt1, pt2);
                                bmp.Save("D:\\" + i.ToString() + "_1.bmp");
                                //计算旋转后的坐标
                                //1.算直线角度
                                double k = 0;
                                if (pt1.X != pt2.X)
                                {
                                    k = (pt2.Y - pt1.Y) / (pt2.X - pt1.X);
                                    jiaodu = Math.Atan(k) / Math.PI * 180;
                                }
                                else { if (pt1.Y < pt2.Y) jiaodu = 90; else jiaodu = -90; }
                                if (jiaodu != 0)
                                {
                                    //如果角度不为零，以图片中心为旋转轴
                                    pt1_r.X = (float)(xyzs[0].X * 25.4) * (float)Math.Cos(jiaodu / 180 * Math.PI) + (float)(xyzs[0].Y * 25.4) * (float)Math.Sin(jiaodu / 180 * Math.PI)+200;
                                    pt1_r.Y = -(float)(xyzs[0].X * 25.4) * (float)Math.Sin(jiaodu / 180 * Math.PI) + (float)(xyzs[0].Y * 25.4) * (float)Math.Cos(jiaodu / 180 * Math.PI)+200;
                                   
                                    //旋转写文字
                                    bmp=KiRotate(bmp, (float)jiaodu, System.Drawing.Color.Transparent);
                                    ht= Graphics.FromImage(bmp);
                                    ht.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;//指定文本呈现的质量 解决文字锯齿问题
                                    ht.DrawString(linelengths[i].ToString("F", m_nfi), myFont, bush, pt1_r);
                                    bmp.Save("D:\\" + i.ToString() + "_2.bmp");
                                    //旋转回来
                                    bmp = KiRotate(bmp, -(float)jiaodu, System.Drawing.Color.Transparent);
                                    ht = Graphics.FromImage(bmp);
                                    bmp.Save("D:\\" + i.ToString() + "_3.bmp");
                                }
                                else {
                                    ht.DrawString(linelengths[i].ToString("F", m_nfi), myFont, bush, pt1);
                                    //ht.DrawString(linelengths[i].ToString("F", m_nfi), myFont, bush, pt2);
                                    bmp.Save("D:\\" + i.ToString() + "_2.bmp");
                                }
                                
                                //
                                i++;

                            }
                            bmp.Save("D:\\00.bmp");
                            
                           
                            
                        }
                    }
                }

                TaskDialog.Show("Revit", info);
                return Autodesk.Revit.UI.Result.Succeeded;

            }
            catch (Exception e)
            {
                message = e.Message;
                return Autodesk.Revit.UI.Result.Failed;
            }
        }

        public static Bitmap KiRotate(Bitmap bmp, float angle, System.Drawing.Color bkColor)
        {
            int w = bmp.Width + 2;
            int h = bmp.Height + 2;

            PixelFormat pf;

            if (bkColor == System.Drawing.Color.Transparent)
            {
                pf = PixelFormat.Format32bppArgb;
            }
            else
            {
                pf = bmp.PixelFormat;
            }

            Bitmap tmp = new Bitmap(w, h, pf);
            Graphics g = Graphics.FromImage(tmp);
            g.Clear(bkColor);
            g.DrawImageUnscaled(bmp, 1, 1);
            g.Dispose();

            GraphicsPath path = new GraphicsPath();
            path.AddRectangle(new RectangleF(0f, 0f, w, h));
            Matrix mtrx = new Matrix();
            mtrx.Rotate(angle);
            RectangleF rct = path.GetBounds(mtrx);

            Bitmap dst = new Bitmap((int)rct.Width, (int)rct.Height, pf);
            g = Graphics.FromImage(dst);
            g.Clear(bkColor);
            g.TranslateTransform(-rct.X, -rct.Y);
            g.RotateTransform(angle);
            g.InterpolationMode = InterpolationMode.HighQualityBilinear;
            g.DrawImageUnscaled(tmp, 0, 0);
            g.Dispose();

            tmp.Dispose();

            return dst;
        }

        public static double ft2m(double val)
        {
            return UnitUtils.ConvertFromInternalUnits(val, DisplayUnitType.DUT_METERS);
        }


        /// <summary>
        /// find out the wall, insert it into a array list
        /// 找出 墙 ，添加到列表
        /// </summary>
        bool initialize()
        {
           ElementSet selections = new ElementSet();
            foreach (ElementId elementId in m_revit.Application.ActiveUIDocument.Selection.GetElementIds())
            {
               selections.Insert(m_revit.Application.ActiveUIDocument.Document.GetElement(elementId));//收集所有选择的ID
            }
            //nothing was selected 什么都没有选择
            if (0 == selections.Size)
            {
                m_errorMessage += "Please select Basic walls";
                return false;
            }

            //find out wall 找出墙
            foreach (Autodesk.Revit.DB.Element e in selections)
            {
                Wall wall = e as Wall;
                if (null != wall)//转换墙 成功
                {
                    if ("Basic" != wall.WallType.Kind.ToString())//进一步判断 是否标准墙  不是则跳过
                    {
                        continue;
                    }
                    m_walls.Add(wall);
                }
            }

            //no wall was selected
            if (0 == m_walls.Count)
            {
                m_errorMessage += "Please select Basic walls";
                return false;
            }
            return true;
        }

        /// <summary>
        /// find out every wall in the selection and add a dimension from the start of the wall to its end
        /// </summary>
        /// <returns>if add successfully, true will be returned, else false will be returned</returns>
        public bool AddDimension()
        {
            if (!initialize())
            {
                return false;
            }

            Transaction transaction = new Transaction(m_revit.Application.ActiveUIDocument.Document, "Add Dimensions");
            transaction.Start();
            //get out all the walls in this array, and create a dimension from its start to its end
            for (int i = 0; i < m_walls.Count; i++)//循环每个墙
            {
                Wall wallTemp = m_walls[i] as Wall;
                if (null == wallTemp)
                {
                    continue;
                }

                //get location curve
                Location      location =     wallTemp.Location;
                LocationCurve locationline = location as LocationCurve;
                if (null == locationline)//判断是否 转换成功
                {
                    continue;//不成功 换下一个
                }

                //New Line

                Line newLine = null;

                //get reference
                ReferenceArray referenceArray = new ReferenceArray();

                AnalyticalModel analyticalModel = wallTemp.GetAnalyticalModel();
                IList<Curve> activeCurveList = analyticalModel.GetCurves(AnalyticalCurveType.ActiveCurves);
                foreach (Curve aCurve in activeCurveList)
                {
                    // find non-vertical curve from analytical model
                    if (aCurve.GetEndPoint(0).Z == aCurve.GetEndPoint(1).Z)
                        newLine = aCurve as Line;
                    if (aCurve.GetEndPoint(0).Z != aCurve.GetEndPoint(1).Z)
                    {
                        AnalyticalModelSelector amSelector = new AnalyticalModelSelector(aCurve);
                        amSelector.CurveSelector = AnalyticalCurveSelector.StartPoint;
                        referenceArray.Append(analyticalModel.GetReference(amSelector));
                    }
                    if (2 == referenceArray.Size)
                        break;
                }
                if (referenceArray.Size != 2)
                {
                    m_errorMessage += "Did not find two references";
                    return false;
                }
                try
                {
                    //try to add new a dimension
                    Autodesk.Revit.UI.UIApplication app = m_revit.Application;
                    Document doc = app.ActiveUIDocument.Document;

                    Autodesk.Revit.DB.XYZ p1 = new XYZ(
                        newLine.GetEndPoint(0).X + 5,
                        newLine.GetEndPoint(0).Y + 5,
                        newLine.GetEndPoint(0).Z);
                    Autodesk.Revit.DB.XYZ p2 = new XYZ(
                        newLine.GetEndPoint(1).X + 5,
                        newLine.GetEndPoint(1).Y + 5,
                        newLine.GetEndPoint(1).Z);

                    Line newLine2 = Line.CreateBound(p1, p2);
                    Dimension newDimension = doc.Create.NewDimension(
                      doc.ActiveView, newLine2, referenceArray);
                }
                // catch the exceptions
                catch (Exception ex)
                {
                    m_errorMessage += ex.ToString();
                    return false;
                }
            }
            transaction.Commit();
            return true;
        }

    }
}
