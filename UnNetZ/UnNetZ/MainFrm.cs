using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace Com.Bing.Frm
{
    public partial class MainFrm : Form
    {
        public MainFrm()
        {
            InitializeComponent();
        }
        private void MainFrm_DragDrop(object sender, DragEventArgs e)
        {
            this.lbFiles.Items.Clear();
            foreach (string path in (IEnumerable<string>)e.Data.GetData(DataFormats.FileDrop))
            {
                if (Path.GetExtension(path).ToLower() == ".exe" && File.Exists(path))
                {
                    this.lbFiles.Items.Add(new ListItem(path));
                }
            }
        }

        private void MainFrm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else e.Effect = DragDropEffects.None;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (lbFiles.Items.Count <= 0)
            {
                MessageBox.Show("请拖入netz组合后的程序", "提示");
                return;
            }
            foreach (ListItem item in lbFiles.Items)
            {
                Dictionary<string, byte[]> resourceDict = UnNetZHelper.GetResourceDict(item.FullPath);
                if (resourceDict == null)
                {
                    MessageBox.Show(string.Format("【{0}】 may not a netZed file.", item.FullPath));
                    continue;
                }
                foreach (KeyValuePair<string, byte[]> pair in resourceDict)
                {
                    if (pair.Key == "zip.dll") continue;
                    using (MemoryStream ms = UnNetZHelper.UnZip(pair.Value))
                    {
                        byte[] bytes = ms.ToArray();
                        try
                        {
                            Assembly assembly = Assembly.Load(bytes);
                            //File.WriteAllBytes(Path.Combine(item.NewFolder, assembly.ManifestModule.ScopeName), bytes);
                            string filePath = Path.Combine(item.NewFolder, assembly.ManifestModule.ScopeName);
                            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                            {
                                fs.Write(bytes, 0, bytes.Length);
                            }
                        }
                        catch
                        {
                            string filePath = Path.Combine(item.NewFolder, UnNetZHelper.UnMangleDllName(pair.Key));
                            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                            {
                                fs.Write(bytes, 0, bytes.Length);
                            }
                        }
                    }
                }
            }
            MessageBox.Show("Done!", "Warning");
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        private const int WM_NCPAINT = 0x0085;
        private const int WM_NCACTIVATE = 0x0086;
        private const int WM_NCLBUTTONDOWN = 0x00A1;
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            Rectangle vRectangle = new Rectangle(3, 3, Width - 6, 21);
            switch (m.Msg)
            {
                case WM_NCPAINT:
                case WM_NCACTIVATE:
                    IntPtr vHandle = GetWindowDC(m.HWnd);
                    using (Graphics vGraphics = Graphics.FromHdc(vHandle))
                    {
                        vGraphics.FillRectangle(new LinearGradientBrush(vRectangle,
                            Color.Pink, Color.Purple, LinearGradientMode.BackwardDiagonal),
                            vRectangle);

                        StringFormat vStringFormat = new StringFormat();
                        vStringFormat.Alignment = StringAlignment.Center;
                        vStringFormat.LineAlignment = StringAlignment.Center;
                        vGraphics.DrawString(this.Text, Font, Brushes.BlanchedAlmond,
                            vRectangle, vStringFormat);
                    }
                    ReleaseDC(m.HWnd, vHandle);
                    break;
                case WM_NCLBUTTONDOWN:
                    Point vPoint = new Point((int)m.LParam);
                    vPoint.Offset(-Left, -Top);
                    if (vRectangle.Contains(vPoint))
                        MessageBox.Show(vPoint.ToString());
                    break;
            }
        }
    }
}