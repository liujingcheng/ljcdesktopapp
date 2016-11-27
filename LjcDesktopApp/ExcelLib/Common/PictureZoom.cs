using System;
using System.Data;
using System.Drawing;
using System.IO;

namespace CanYouLib.ExcelLib.Utility
{
    /// <summary>
    /// ͼƬ���š�Code by:�����
    /// </summary>
    public  class PictureZoom
    {
        /// <summary>
        /// ��СͼƬ
        /// </summary>
        /// <param name="strOldPic">Դͼ�ļ���(����·��)</param>
        /// <param name="strNewPic">��С�󱣴�Ϊ�ļ���(����·��)</param>
        /// <param name="intWidth">��С�����</param>
        /// <param name="intHeight">��С���߶�</param>
        public void Zoom(string strOldPic, string strNewPic, int intWidth, int intHeight)
        {
            if (intWidth > intHeight)//�������ڸߣ���ָ���ߣ��Զ������
                ZoomAutoWith(strOldPic, strNewPic, intHeight);
            else
                ZoomAutoHight(strOldPic, strNewPic, intWidth);
        }

        /// <summary>
        /// ��������СͼƬ���Զ�����߶�
        /// </summary>
        /// <param name="strOldPic">Դͼ�ļ���(����·��)</param>
        /// <param name="strNewPic">��С�󱣴�Ϊ�ļ���(����·��)</param>
        /// <param name="intWidth">��С�����</param>
        public void ZoomAutoHight(string strOldPic, string strNewPic, int intWidth)
        {

            System.Drawing.Bitmap objPic, objNewPic;
            try
            {
                objPic = new System.Drawing.Bitmap(strOldPic);
                //int intHeight = (intWidth * objPic.Height) / objPic.Width;//�Զ�����߶ȣ���ԭ��������
                //objNewPic = new System.Drawing.Bitmap(objPic, intWidth, intHeight);
                objNewPic = BitmapZoomAutoHight(objPic, intWidth);
                objNewPic.Save(strNewPic,System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            catch (Exception exp) { throw exp; }
            finally
            {
                objPic = null;
                objNewPic = null;
            }
        }
        /// <summary>
        /// ��������СͼƬ���Զ�������
        /// </summary>
        /// <param name="strOldPic">Դͼ�ļ���(����·��)</param>
        /// <param name="strNewPic">��С�󱣴�Ϊ�ļ���(����·��)</param>
        /// <param name="intHeight">��С���߶�</param>
        public void ZoomAutoWith(string strOldPic, string strNewPic, int intHeight)
        {

            System.Drawing.Bitmap objPic, objNewPic;
            try
            {
                objPic = new System.Drawing.Bitmap(strOldPic);
                //int intWidth = (intHeight * objPic.Width) / objPic.Height;//�Զ������ȣ���ԭ��������
                //objNewPic = new System.Drawing.Bitmap(objPic, intWidth, intHeight);
                objNewPic = BitmapZoomAutoWith(objPic, intHeight);
                objNewPic.Save(strNewPic, System.Drawing.Imaging.ImageFormat.Jpeg);

            }
            catch (Exception exp) { throw exp; }
            finally
            {
                objPic = null;
                objNewPic = null;
            }
        }

        /// <summary>
        /// ��Bitmap���ݿ���Զ�����
        /// </summary>
        /// <param name="p_Bitmap">ͼƬԴ</param>
        /// <param name="intWidth">���</param>
        /// <returns></returns>
        private System.Drawing.Bitmap BitmapZoomAutoHight(Bitmap p_Bitmap, int intWidth)
        {
            int intHeight = (intWidth * p_Bitmap.Height) / p_Bitmap.Width;//�Զ�����߶ȣ���ԭ��������
            return new System.Drawing.Bitmap(p_Bitmap, intWidth, intHeight);
        }
        /// <summary>
        /// ��Bitmap���ݿ���Զ�����
        /// </summary>
        /// <param name="p_Bitmap">ͼƬԴ</param>
        /// <param name="intHeight">�߶�</param>
        /// <returns></returns>
        private System.Drawing.Bitmap BitmapZoomAutoWith(Bitmap p_Bitmap, int intHeight)
        {
            int intWidth = (intHeight * p_Bitmap.Width) / p_Bitmap.Height;//�Զ������ȣ���ԭ��������
            return new System.Drawing.Bitmap(p_Bitmap, intWidth, intHeight);
        }

        /// <summary>
        /// ����Bitmap
        /// </summary>
        /// <param name="p_MaxHeight">�����ø߶Ƚ�������</param>
        /// <param name="p_MaxWidth">�����ÿ�Ƚ�������</param>
        /// <param name="p_Bitmap">ͼƬԴ</param>
        /// <returns></returns>
        public Bitmap ZoomBitmap(int p_MaxHeight, int p_MaxWidth, Bitmap p_Bitmap)
        {
            //ת����bmp���ͼƬ��ԭʼ��С�뵥Ԫ��Ĵ�С�Ƚϣ������Ƿ���Ҫ��������
            if (p_Bitmap.Size.Height > p_MaxHeight || p_Bitmap.Size.Width > p_MaxWidth)//��Ҫ����
            {
                //���ݵ�Ԫ���������������Զ����ſ���

                if (p_MaxHeight > p_MaxWidth)//Ҳ����˵�����ָ�������߾����Զ�����
                    return BitmapZoomAutoHight(p_Bitmap, p_MaxWidth);
                else
                    return BitmapZoomAutoWith(p_Bitmap, p_MaxHeight);
            }
            else
                return new Bitmap(p_Bitmap, p_Bitmap.Size.Width, p_Bitmap.Size.Height);
        }
        /// <summary>
        /// ����Bitmap
        /// </summary>
        /// <param name="p_MaxHeight">�����ø߶Ƚ�������</param>
        /// <param name="p_MaxWidth">�����ÿ�Ƚ�������</param>
        /// <param name="p_Bitmap">ͼƬԴ</param>
        /// <param name="p_FileNameToSave">ѹ������ļ�����·��</param>
        /// <returns></returns>
        public void ZoomBitmap(int p_MaxHeight, int p_MaxWidth, Bitmap p_Bitmap, string p_FileNameToSave)
        {
            ZoomBitmap(p_MaxHeight, p_MaxWidth, p_Bitmap).Save(p_FileNameToSave, System.Drawing.Imaging.ImageFormat.Jpeg);
        }


    }
}
