using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Drawing;
using CanYouLib.ExcelLib.Utility;

/***
 * Create Version:0.0.0.2
 * Code by:�����
 * Date:2010-04-08
 * ���¼�¼��
 * Code by:�����
 * Date:2011-03-11
 * 1�������˲����в���������ǰһ����ʽ��
 * 2�������˲���ͼƬ���ܣ�����֧���Զ����š�
 * 3��֧����ҳüҳ�ŵ�����
 * Date:2011-06-07
 * 1��֧��ģ����ͷ��������
 * 2�����Ӻϲ��й���
 * Date:2011-09-13
 * 1�������˹�ʽ֧��
 * 2��������_workbook��_sheet����
 ***/
namespace CanYouLib.ExcelLib
{
    /// <summary>
    /// ����Excelģ��ı��������ࡣ
    /// ����� 2010-04-08
    /// ע����������Office.Excel
    /// </summary>
    public class ExcelExpert : IDisposable
    {
        private IWorkbook _workbook;
        private ISheet _sheet;
        private string _templateFileName;
        /// <summary>
        /// ģ���ļ��������ļ���
        /// </summary>
        public string TemplateFileName
        {
            get { return _templateFileName; }
            set { _templateFileName = value; }
        }

        /// <summary>
        /// ���캯��,Ĭ�ϲ�����һ��Sheet��
        /// </summary>
        /// <param name="p_FileName">ģ���ļ�����·��</param>
        public ExcelExpert(string p_FileName)
        {

            this.TemplateFileName = p_FileName;
            //��Excel�ļ����������ݷŵ�HSSFWorkbook����
            using (var file = new FileStream(p_FileName, FileMode.Open, FileAccess.Read))
            {
                _workbook = WorkbookFactory.Create(file);
            }
            //���õ�ǰ������Sheet��Ĭ��Ϊ��һ���������Ϳ��Բ��������
            _sheet = _workbook.GetSheetAt(0);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_FileName"></param>
        /// <param name="p_SheetIndex"></param>
        public ExcelExpert(string p_FileName, int p_SheetIndex)
        {

            this.TemplateFileName = p_FileName;
            //ɾ����ʱ�ļ�

            //��Excel�ļ����������ݷŵ�HSSFWorkbook����
            using (FileStream file = new FileStream(p_FileName, FileMode.Open, FileAccess.Read))
            {
                _workbook = WorkbookFactory.Create(file);
            }
            //���õ�ǰ������Sheet��Ĭ��Ϊ��һ���������Ϳ��Բ��������
            _sheet = _workbook.GetSheetAt(p_SheetIndex);
        }

        #region SetValue �Ե�Ԫ������ֵ public
        /// <summary>
        /// �Ե�Ԫ������ֵ
        /// </summary>
        /// <param name="p_startRowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_endRowIndex">�����кţ���1��ʼ</param>
        /// <param name="p_startColChars">��ʼ��������A��ʼ</param>
        /// <param name="p_endColChars">������������A��ʼ</param>
        /// <param name="p_Values"></param>
        public void SetValue(int p_startRowIndex, int p_endRowIndex, string p_startColChars, string p_endColChars, Array p_Values)
        {
            SetValue(p_startRowIndex, p_endRowIndex, ExcelColumnTranslator.ToIndex(p_startColChars), ExcelColumnTranslator.ToIndex(p_endColChars), p_Values);
        }
        /// <summary>
        /// �Ե�Ԫ������ֵ
        /// </summary>
        /// <param name="p_startRowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_endRowIndex">�����кţ���1��ʼ</param>
        /// <param name="p_startColIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_endColIndex">�����кţ���1��ʼ</param>
        /// <param name="p_Values"></param>
        public void SetValue(int p_startRowIndex, int p_endRowIndex, int p_startColIndex, int p_endColIndex, Array p_Values)
        {
            //����������
            int index1 = 0;
            int index2 = 0;
            for (int rowIndex = p_startRowIndex; rowIndex <= p_endRowIndex; rowIndex++)
            {
                index2 = 0;
                for (int colIndex = p_startColIndex; colIndex <= p_endColIndex; colIndex++)
                {
                    SetValue(rowIndex, colIndex, p_Values.GetValue(index1, index2).ToString());
                    index2++;
                }
                index1++;
            }
        }



        /// <summary>
        /// �Ե�Ԫ��ֵ
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColChars">��������A��ʼ</param>
        /// <param name="p_Value">����ֵ</param>
        public void SetValue(int p_RowIndex, string p_ColChars, string p_Value)
        {
            SetValue(p_RowIndex, ExcelColumnTranslator.ToIndex(p_ColChars), p_Value);
        }

        /// <summary>
        /// �Ե�Ԫ��ֵ
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColChars">��������A��ʼ</param>
        /// <param name="p_Value">����ֵ</param>
        public void SetValue(int p_RowIndex, string p_ColChars, double p_Value)
        {
            SetValue(p_RowIndex, ExcelColumnTranslator.ToIndex(p_ColChars), p_Value);
        }



        /// <summary>
        /// �Ե�Ԫ��ֵ
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColChars">��������A��ʼ</param>
        /// <param name="p_Value">����ֵ</param>
        public void SetValue(int p_RowIndex, string p_ColChars, DateTime p_Value)
        {
            SetValue(p_RowIndex, ExcelColumnTranslator.ToIndex(p_ColChars), p_Value);
        }





        /// <summary>
        /// �Ե�Ԫ��ֵ
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColumnIndex">�кţ���1��ʼ</param>
        /// <param name="p_Value">����ֵ</param>
        public void SetValue(int p_RowIndex, int p_ColumnIndex, DateTime p_Value)
        {
            _sheet.GetRow(p_RowIndex - 1).GetCell(p_ColumnIndex - 1).SetCellValue(p_Value);
        }

        /// <summary>
        /// �Ե�Ԫ��ֵ
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColumnIndex">�кţ���1��ʼ</param>
        /// <param name="p_Value">����ֵ����1��ʼ</param>
        private void SetValue(int p_RowIndex, int p_ColumnIndex, string p_Value)
        {
            _sheet.GetRow(p_RowIndex - 1).GetCell(p_ColumnIndex - 1).SetCellValue(p_Value);
        }
        /// <summary>
        /// �Ե�Ԫ��ֵ
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColumnIndex">�кţ���1��ʼ</param>
        /// <param name="p_Value">����ֵ</param>
        public void SetValue(int p_RowIndex, int p_ColumnIndex, double p_Value)
        {
            _sheet.GetRow(p_RowIndex - 1).GetCell(p_ColumnIndex - 1).SetCellValue(p_Value);
        }
        #endregion

        #region SetColumnHidden ������ public
        /// <summary>
        /// ����������
        /// </summary>
        /// <param name="p_StartcolIndex">���ص���ʼ�����</param>
        /// <param name="p_EndcolIndex">���ص���ֹ�����</param>
        public void SetColumnHidden(int p_StartcolIndex, int p_EndcolIndex)
        {
            for (int i = p_StartcolIndex; i <= p_EndcolIndex; i++)
            {
                _sheet.SetColumnHidden(i, true);
            }
        }


        /// <summary>
        /// ����������
        /// </summary>
        /// <param name="p_StartcolChar">���ص���ʼ�����</param>
        /// <param name="p_EndcolChar">���ص���ֹ�����</param>
        public void SetColumnHidden(string p_StartcolChar, string p_EndcolChar)
        {
            SetColumnHidden(ExcelColumnTranslator.ToIndex(p_StartcolChar) - 1,
                ExcelColumnTranslator.ToIndex(p_EndcolChar) - 1);
        }
        #endregion

        #region AddMergedRegion �ϲ���Ԫ�� public
        /// <summary>
        /// ��Ӻϲ������罫C3:E5�ϲ�Ϊһ����Ԫ��,��ô��������Ϊ��
        /// AddMergedRegion(3,5,"C","E")
        /// </summary>
        /// <param name="p_StartRowIndex">��ʼ��</param>
        /// <param name="p_EndRowIndex">������</param>
        /// <param name="p_StartColChars">��ʼ�б�ʶ</param>
        /// <param name="p_EndColChars">�����б�ʶ</param>
        public void AddMergedRegion(int p_StartRowIndex, int p_EndRowIndex, string p_StartColChars, string p_EndColChars)
        {
            _sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(p_StartRowIndex - 1, ExcelColumnTranslator.ToIndex(p_StartColChars) - 1, p_EndRowIndex - 1, ExcelColumnTranslator.ToIndex(p_EndColChars) - 1));
        }
        #endregion

        #region InsertPicture ��ָ����Ԫ�����ͼƬ public
        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_RowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_ColIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_FileName">ͼƬԴ�ļ�����·��</param>
        public void InsertPicture(int p_RowIndex, int p_ColIndex, string p_FileName)
        {
            InsertPicture(p_RowIndex, p_RowIndex, p_ColIndex, p_ColIndex, p_FileName, 10, 10);
        }

        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_RowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_ColIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_FileName">ͼƬԴ�ļ�����·��</param>
        /// <param name="p_LeftRectify">ͼƬ��Ե�Ԫ�����ƫ��ֵ�����ֵ1023</param>
        /// <param name="p_TopRectify">ͼƬ��Ե�Ԫ���ϵ�ƫ��ֵ�����ֵ255</param>
        public void InsertPicture(int p_RowIndex, int p_ColIndex, string p_FileName, int p_LeftRectify, int p_TopRectify)
        {
            InsertPicture(p_RowIndex, p_RowIndex, p_ColIndex, p_ColIndex, p_FileName, p_LeftRectify, p_TopRectify);
        }

        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_RowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_ColChars">��ʼ��������A��ʼ</param>
        /// <param name="p_FileName">ͼƬԴ�ļ�����·��</param>
        public void InsertPicture(int p_RowIndex, string p_ColChars, string p_FileName)
        {
            InsertPicture(p_RowIndex, p_RowIndex, ExcelColumnTranslator.ToIndex(p_ColChars), ExcelColumnTranslator.ToIndex(p_ColChars), p_FileName, 10, 10);
        }

        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_RowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_ColChars">��ʼ��������A��ʼ</param>
        /// <param name="p_FileName">ͼƬԴ�ļ�����·��</param>
        /// <param name="p_LeftRectify">ͼƬ��Ե�Ԫ�����ƫ��ֵ�����ֵ1023</param>
        /// <param name="p_TopRectify">ͼƬ��Ե�Ԫ���ϵ�ƫ��ֵ�����ֵ255</param>
        public void InsertPicture(int p_RowIndex, string p_ColChars, string p_FileName, int p_LeftRectify, int p_TopRectify)
        {
            InsertPicture(p_RowIndex, p_RowIndex, ExcelColumnTranslator.ToIndex(p_ColChars), ExcelColumnTranslator.ToIndex(p_ColChars), p_FileName, p_LeftRectify, p_TopRectify);
        }


        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_StartRowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_EndRowIndex">�����кţ���1��ʼ</param>
        /// <param name="p_StartColChars">��ʼ��������A��ʼ</param>
        /// <param name="p_EndColChars">������������A��ʼ</param>
        /// <param name="p_FileName">ͼƬԴ�ļ�����·��</param>
        public void InsertPicture(int p_StartRowIndex, int p_EndRowIndex, string p_StartColChars, string p_EndColChars, string p_FileName)
        {
            InsertPicture(p_StartRowIndex, p_EndRowIndex, ExcelColumnTranslator.ToIndex(p_StartColChars), ExcelColumnTranslator.ToIndex(p_EndColChars), p_FileName, 10, 10);
        }


        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_StartRowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_EndRowIndex">�����кţ���1��ʼ</param>
        /// <param name="p_StartColChars">��ʼ��������A��ʼ</param>
        /// <param name="p_EndColChars">������������A��ʼ</param>
        /// <param name="p_FileName">ͼƬԴ�ļ�����·��</param>
        /// <param name="p_LeftRectify">ͼƬ��Ե�Ԫ�����ƫ��ֵ�����ֵ1023</param>
        /// <param name="p_TopRectify">ͼƬ��Ե�Ԫ���ϵ�ƫ��ֵ�����ֵ255</param>
        public void InsertPicture(int p_StartRowIndex, int p_EndRowIndex, string p_StartColChars, string p_EndColChars, string p_FileName, int p_LeftRectify, int p_TopRectify)
        {
            InsertPicture(p_StartRowIndex, p_EndRowIndex, ExcelColumnTranslator.ToIndex(p_StartColChars), ExcelColumnTranslator.ToIndex(p_EndColChars), p_FileName, p_LeftRectify, p_TopRectify);
        }

        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_StartRowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_EndRowIndex">�����кţ���1��ʼ</param>
        /// <param name="p_StartColumnIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_EndColumnIndex">�����кţ���1��ʼ</param>
        /// <param name="p_FileName">ͼƬԴ�ļ�����·��</param>
        public void InsertPicture(int p_StartRowIndex, int p_EndRowIndex, int p_StartColumnIndex, int p_EndColumnIndex, string p_FileName)
        {
            InsertPicture(p_StartRowIndex, p_EndRowIndex, p_StartColumnIndex, p_EndColumnIndex, p_FileName, 10, 10);
        }

        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColChars">��������A��ʼ</param>
        /// <param name="p_Bitmap">ͼƬԴ</param>
        public void InsertPicture(int p_RowIndex, string p_ColChars, Bitmap p_Bitmap)
        {
            InsertPicture(p_RowIndex, p_RowIndex, ExcelColumnTranslator.ToIndex(p_ColChars), ExcelColumnTranslator.ToIndex(p_ColChars), p_Bitmap, 10, 10);
        }

        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColChars">��������A��ʼ</param>
        /// <param name="p_Bitmap">ͼƬԴ</param>
        /// <param name="p_LeftRectify">ͼƬ��Ե�Ԫ�����ƫ��ֵ�����ֵ1023</param>
        /// <param name="p_TopRectify">ͼƬ��Ե�Ԫ���ϵ�ƫ��ֵ�����ֵ255</param>
        public void InsertPicture(int p_RowIndex, string p_ColChars, Bitmap p_Bitmap, int p_LeftRectify, int p_TopRectify)
        {
            InsertPicture(p_RowIndex, p_RowIndex, ExcelColumnTranslator.ToIndex(p_ColChars), ExcelColumnTranslator.ToIndex(p_ColChars), p_Bitmap, p_LeftRectify, p_TopRectify);
        }

        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_StartRowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_EndRowIndex">�����кţ���1��ʼ</param>
        /// <param name="p_StartColChars">��ʼ��������A��ʼ</param>
        /// <param name="p_EndColChars">������������A��ʼ</param>
        /// <param name="p_Bitmap">ͼƬԴ</param>
        /// <param name="p_LeftRectify">ͼƬ��Ե�Ԫ�����ƫ��ֵ�����ֵ1023</param>
        /// <param name="p_TopRectify">ͼƬ��Ե�Ԫ���ϵ�ƫ��ֵ�����ֵ255</param>
        public void InsertPicture(int p_StartRowIndex, int p_EndRowIndex, string p_StartColChars, string p_EndColChars, Bitmap p_Bitmap, int p_LeftRectify, int p_TopRectify)
        {
            InsertPicture(p_StartRowIndex, p_EndRowIndex, ExcelColumnTranslator.ToIndex(p_StartColChars), ExcelColumnTranslator.ToIndex(p_EndColChars), p_Bitmap, p_LeftRectify, p_TopRectify);
        }

        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_StartRowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_EndRowIndex">�����кţ���1��ʼ</param>
        /// <param name="p_StartColumnIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_EndColumnIndex">�����кţ���1��ʼ</param>
        /// <param name="p_Bitmap">ͼƬԴ</param>
        /// <param name="p_LeftRectify">ͼƬ��Ե�Ԫ�����ƫ��ֵ�����ֵ1023</param>
        /// <param name="p_TopRectify">ͼƬ��Ե�Ԫ���ϵ�ƫ��ֵ�����ֵ255</param>
        public void InsertPicture(int p_StartRowIndex, int p_EndRowIndex, int p_StartColumnIndex, int p_EndColumnIndex, Bitmap p_Bitmap, int p_LeftRectify, int p_TopRectify)
        {
            p_StartRowIndex = p_StartRowIndex - 1;
            p_EndRowIndex = p_EndRowIndex - 1;
            p_StartColumnIndex = p_StartColumnIndex - 1;
            p_EndColumnIndex = p_EndColumnIndex - 1;

            //����
            p_Bitmap = PictureZoom(p_StartRowIndex, p_EndRowIndex, p_StartColumnIndex, p_EndColumnIndex, p_Bitmap);

            //��ͼƬ������ӵ�workbook��
            int pictureIdx = _workbook.AddPicture(FileUtility.ImageToByte(p_Bitmap), PictureType.JPEG);

            //����һ�������ͼ�λ滭�ߡ�
            IDrawing patriarch = _sheet.CreateDrawingPatriarch();
            var anchor = new HSSFClientAnchor
            {
                Col1 = p_StartColumnIndex,
                Row1 = p_StartRowIndex,
                Col2 = p_EndColumnIndex + 1,
                Row2 = p_EndRowIndex + 1
            };
            //��ʾͼƬ
            var pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
            //ȡ���Զ�����Ч������ʱanchor����ƫ������Ч��Ҫ��������
            pict.Resize();
            pict.Anchor.Dx1 = p_LeftRectify;
            pict.Anchor.Dy1 = p_TopRectify;
        }


        /// <summary>
        /// ��ͼƬ����Ԫ�����򣬸��ݵ�Ԫ�������Զ�����
        /// </summary>
        /// <param name="p_StartRowIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_EndRowIndex">�����кţ���1��ʼ</param>
        /// <param name="p_StartColumnIndex">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_EndColumnIndex">�����кţ���1��ʼ</param>
        /// <param name="p_LeftRectify">ͼƬ��Ե�Ԫ�����ƫ��ֵ�����ֵ1023</param>
        /// <param name="p_TopRectify">ͼƬ��Ե�Ԫ���ϵ�ƫ��ֵ�����ֵ255</param>
        /// <param name="p_FileName">ͼƬԴ�ļ�����·��</param>
        public void InsertPicture(int p_StartRowIndex, int p_EndRowIndex, int p_StartColumnIndex, int p_EndColumnIndex, string p_FileName, int p_LeftRectify, int p_TopRectify)
        {
            //���ļ��浽Bitmap
            var bitmap = new Bitmap(p_FileName);

            //����
            InsertPicture(p_StartRowIndex, p_EndRowIndex, p_StartColumnIndex, p_EndColumnIndex, bitmap, p_LeftRectify, p_TopRectify);
        }
        #endregion

        #region SetFormula  �Ե�Ԫ�����ù�ʽ public
        /// <summary>
        /// �Ե�Ԫ�����ù�ʽ
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColChars">��������A��ʼ</param>
        /// <param name="p_Formula">��ʽ���磺A1*A2</param>
        public void SetFormula(int p_RowIndex, string p_ColChars, string p_Formula)
        {
            SetFormula(p_RowIndex, ExcelColumnTranslator.ToIndex(p_ColChars), p_Formula);
        }
        /// <summary>
        /// �Ե�Ԫ�����ù�ʽ
        /// </summary>
        /// <param name="p_RowIndex">�кţ���1��ʼ</param>
        /// <param name="p_ColumnIndex">�кţ���1��ʼ</param>
        /// <param name="p_Value">��ʽ���磺A1*A2</param>
        public void SetFormula(int p_RowIndex, int p_ColumnIndex, string p_Formula)
        {
            _sheet.GetRow(p_RowIndex - 1).GetCell(p_ColumnIndex - 1).SetCellFormula(p_Formula);
        }
        #endregion

        #region CreateRows ����ʼ���Ϸ���������,����ֵ��ʼ�е���ʽ public
        /// <summary>
        /// ����ʼ���Ϸ���������,����ֵ��ʼ�е���ʽ��
        /// </summary>
        /// <param name="p_StartRow">��ʼ�кţ���1��ʼ</param>
        /// <param name="p_Count">��Ҫ����������</param>
        public void CreateRows(int p_StartRow, int p_Count)
        {
            p_Count = p_Count - 1;
            p_StartRow = p_StartRow - 1;
            IRow mySourceStyleRow = this._sheet.GetRow(p_StartRow);//��ȡԴ��ʽ��
            //_sheet.ShiftRows(p_StartRow + 1, this._sheet.LastRowNum, p_Count, true, false, true);
            ////���ò����з���

            MyInsertRow(this._sheet, p_StartRow, p_Count, mySourceStyleRow);

        }
        #endregion

        #region ExcelToByte  ת����ʽ���÷������Զ���ת����ɾ��ԭ�ļ� public
        /// <summary>
        /// ת����ʽ���÷������Զ���ת����ɾ��ԭ�ļ���
        /// </summary>
        /// <param name="p_FileName"></param>
        /// <returns></returns>
        public byte[] ExcelToByte(string p_FileName)
        {
            try
            {
                byte[] byteResult = ReadFile(p_FileName);

                File.Delete(p_FileName);


                return byteResult;
            }
            catch (Exception e)
            {

                if (File.Exists(p_FileName))
                    File.Delete(p_FileName);
                throw new Exception("��Excel�ļ�תΪByteʱ����������Ϣ��" + e.Message);
            }
        }
        #endregion

        #region �����ļ�
        /// <summary>
        /// �����ļ�
        /// </summary>
        /// <param name="p_FileName"></param>
        public void Save(string p_FileName)
        {
            using (var file = new FileStream(p_FileName, FileMode.Create))
            {
                _workbook.Write(file);
                file.Close();
            }

        }
        #endregion

        #region �������� private
        /// <summary>
        /// ��Bitmap���ݵ�Ԫ�������С��������
        /// </summary>
        /// <param name="p_StartRowIndex"></param>
        /// <param name="p_EndRowIndex"></param>
        /// <param name="p_StartColumnIndex"></param>
        /// <param name="p_EndColumnIndex"></param>
        /// <param name="p_Bitmap"></param>
        /// <returns></returns>
        private Bitmap PictureZoom(int p_StartRowIndex, int p_EndRowIndex, int p_StartColumnIndex, int p_EndColumnIndex, Bitmap p_Bitmap)
        {
            //����ȡ��Excel��Ԫ�����С
            //�и�/15=����
            int height = 0;
            for (int j = p_StartRowIndex; j <= p_EndRowIndex; j++)
                height += (int)(_sheet.GetRow(j).Height / 15);
            int with = 0;
            //�п�/32=����
            for (int i = p_StartColumnIndex; i <= p_EndColumnIndex; i++)
                with += (int)(_sheet.GetColumnWidth(i) / 32);

            //�����������
            var pz = new PictureZoom();

            return pz.ZoomBitmap(height, with, p_Bitmap);
        }


        /// <summary>
        /// ��ȡ�ļ���
        /// </summary>
        /// <param name="p_FileName"></param>
        /// <returns></returns>
        private byte[] ReadFile(string p_FileName)
        {
            FileStream pFileStream = null;

            var pReadByte = new byte[0];

            try
            {

                pFileStream = new FileStream(p_FileName, FileMode.Open, FileAccess.Read);

                var r = new BinaryReader(pFileStream);

                r.BaseStream.Seek(0, SeekOrigin.Begin);    //���ļ�ָ�����õ��ļ���

                pReadByte = r.ReadBytes((int)r.BaseStream.Length);
                pFileStream.Close();
                pFileStream = null;
                return pReadByte;

            }

            catch (Exception e)
            {
                if (pFileStream != null)
                    pFileStream.Close();
                throw new Exception("��ȡ��������ļ�����������Ϣ" + e.Message);
            }

        }

        /// <summary>
        /// ����ʼ���Ϸ���������,����ֵ��ʼ�е���ʽ��
        /// </summary>
        /// <param name="sheet">sheet</param>
        /// <param name="InsertRow">������</param>
        /// <param name="InsertCountRows">����������</param>
        /// <param name="RowStyles">��ʽ</param>
        private void MyInsertRow(ISheet sheet, int InsertRow, int InsertCountRows, IRow RowStyles)
        {

            #region �������ƶ���ճ��Ŀ��в壬������Ӧ���У����Բ����е���һ��Ϊ��ʽԴ(����������-1����һ��)
            for (int i = InsertRow; i <= InsertRow + InsertCountRows - 1; i++)
            {
                IRow targetRow = null;
                ICell sourceCell = null;
                ICell targetCell = null;

                targetRow = sheet.CreateRow(i + 1);

                for (int m = RowStyles.FirstCellNum; m < RowStyles.LastCellNum; m++)
                {
                    sourceCell = RowStyles.GetCell(m);
                    if (sourceCell == null)
                        continue;
                    targetCell = targetRow.CreateCell(m);
                    targetCell.CellStyle = sourceCell.CellStyle;
                    targetCell.SetCellType(sourceCell.CellType);

                }
            }
            #endregion
        }
        #endregion
        /// <summary>
        /// �ͷ���Դ
        /// </summary>
        void IDisposable.Dispose()
        {
            _sheet = null;
            _workbook = null;
        }

    }
}
