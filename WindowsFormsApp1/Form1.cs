using System;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.Globalization;
//using Excel = Microsoft.Office.Interop.Excel;
//using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
    public partial class MNLZ : Form
    {

        public MNLZ()
        {
            InitializeComponent();
        }

        //   public event ErrorEventHandler RaiseCustomEvent;

        struct MNLZZone  //Структура
        {
            public float Length, Tau1, Tau2, Alpha, Etha, Phi, For1, gF, G, GCurr, DeltaG;
        };

        float Lzvo, Tmet, Tvux, Bsl, Csl;
        static int k = 1;
        static int n = 1;
        static int p = 0;
        static int mintimer = 0;
        static int timer = 0;

        float[] Tz = new float[11];
        float[] tpov = new float[11];
        float[] Tz1 = new float[11];
        float[] EZ = new float[11];
        float[] E = new float[11];

        float[] Koeff = new float[11];
        float[] Y = new float[5001];
        float[] T = new float[11];
        float[] Y1 = new float[5001];
        float[] E11 = new float[11];

        MNLZZone[] ZVO = new MNLZZone[11]; //массив структур

        public static float Alpha0(float Tau)
        {
            float Result = 0;

            if ((Tau >= 0) && (Tau <= 2 * 60))
                Result = 350;

            if ((Tau > 2 * 60) && (Tau <= 4 * 60))
                Result = 270;

            if ((Tau > 4 * 60) && (Tau <= 6 * 60))
                Result = 250;

            if (Tau > 6 * 60)
                Result = 210;

            return Result;
        }

        private void button1_Click(object sender, EventArgs e) //Исходные данные
        {
            textBox1.Text = "0,24"; //zona 1
            textBox2.Text = "0,57"; //zona 2
            textBox3.Text = "0,94"; //zona 3
            textBox4.Text = "1,36"; //zona 4
            textBox5.Text = "1,92"; //zona 5
            textBox6.Text = "3,84"; //zona 6
            textBox7.Text = "3,88"; //zona 7 
            textBox8.Text = "4,73"; //zona 8
            textBox9.Text = "9,57"; //zona 9
            textBox10.Text = "0,9"; //Толщина слитка
            textBox11.Text = "1,5"; //Ширина слитка
            textBox12.Text = "0,20"; //Толщина слитка
            textBox13.Text = "1506"; //Температруа при  заливке
            textBox14.Text = "600"; //Температруа на выходе
        }

        private void button3_Click(object sender, EventArgs e) //Отчистка полей
        {
            timer1.Enabled = false;
            timer1.Stop();
            timer2.Enabled = false;
            timer2.Stop();

            textBox16.BackColor = Color.Red;

            hScrollBar1.Value = 10;
            trackBar1.Value = 0;
            label84.Text = trackBar1.Value.ToString() + " м/мин";

            this.chart1.Series[0].Points.Clear();
            this.chart1.Series[1].Points.Clear();
            this.chart1.Series[2].Points.Clear();

            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            dataGridView1.Rows.Clear();

            label41.Text = "0,00 м3/год";
            label42.Text = "0,00 м3/год";
            label43.Text = "0,00 м3/год";
            label44.Text = "0,00 м3/год";
            label45.Text = "0,00 м3/год";
            label46.Text = "0,00 м3/год";
            label47.Text = "0,00 м3/год";
            label48.Text = "0,00 м3/год";
            label49.Text = "0,00 м3/год";

            label59.Text = "0,00 м3/год";
            label60.Text = "0,00 м3/год";
            label61.Text = "0,00 м3/год";
            label62.Text = "0,00 м3/год";
            label63.Text = "0,00 м3/год";
            label64.Text = "0,00 м3/год";
            label65.Text = "0,00 м3/год";
            label66.Text = "0,00 м3/год";

            label79.Text = "0,00 м3/год";
            label80.Text = "0,00 м3/год";

            label83.Text = "0,00 хв";
            label89.Text = "0,00 сек";

            label85.Text = "0,00 м3/год";
            label86.Text = "0,00 °С";

            label40.Text = "0,00 °С";
            label50.Text = "0,00 °С";
            label51.Text = "0,00 °С";
            label52.Text = "0,00 °С";
            label53.Text = "0,00 °С";
            label54.Text = "0,00 °С";
            label55.Text = "0,00 °С";
            label56.Text = "0,00 °С";
            label57.Text = "0,00 °С";
        }

        private void button2_Click(object sender, EventArgs e) //Расчет
        {


            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox7.Text != "" && textBox8.Text != "" && textBox9.Text != "" && textBox10.Text != "" && textBox11.Text != "" && textBox12.Text != "" && textBox13.Text != "" && textBox14.Text != "")
            {
                if (textBox1.Text != "0" && textBox2.Text != "0" && textBox3.Text != "0" && textBox4.Text != "0" && textBox5.Text != "0" && textBox6.Text != "0" && textBox7.Text != "0" && textBox8.Text != "0" && textBox9.Text != "0" && textBox10.Text != "0" && textBox11.Text != "0" && textBox12.Text != "0" && textBox13.Text != "0" && textBox14.Text != "0")
                {
                    if (textBox1.Text != "," && textBox2.Text != "," && textBox3.Text != "," && textBox4.Text != "," && textBox5.Text != "," && textBox6.Text != "," && textBox7.Text != "," && textBox8.Text != "," && textBox9.Text != "," && textBox10.Text != "," && textBox11.Text != "," && textBox12.Text != "," && textBox13.Text != "," && textBox14.Text != ",")
                    {
                        string z1 = textBox1.Text; //Зона 1
                        string z2 = textBox2.Text; //Зона 2
                        string z3 = textBox3.Text; //Зона 3
                        string z4 = textBox4.Text; //Зона 4
                        string z5 = textBox5.Text; //Зона 5
                        string z6 = textBox6.Text; //Зона 6
                        string z7 = textBox7.Text; //Зона 7
                        string z8 = textBox8.Text; //Зона 8
                        string z9 = textBox9.Text; //Зона 9

                        // float z_1 = Convert.ToSingle(z1);
                        float z_1 = Convert.ToSingle(z1);
                        float z_2 = Convert.ToSingle(z2);
                        float z_3 = Convert.ToSingle(z3);
                        float z_4 = Convert.ToSingle(z4);
                        float z_5 = Convert.ToSingle(z5);
                        float z_6 = Convert.ToSingle(z6);
                        float z_7 = Convert.ToSingle(z7);
                        float z_8 = Convert.ToSingle(z8);
                        float z_9 = Convert.ToSingle(z9);

                        textBox16.BackColor = Color.Green;

                        for (int i = 1; i <= 10; i++) //for (int i = 1; i <= 10; i++)
                            ZVO[i].GCurr = 0;

                        if (z_1 > 0.1 && z_2 > 0.1 && z_3 > 0.1 && z_4 > 0.1 && z_5 > 0.1 && z_6 > 0.1 && z_7 > 0.1 && z_8 > 0.1 && z_9 > 0.1)
                        {
                            string l11 = textBox11.Text; //Ширина слитка
                            string l12 = textBox12.Text; //Высота слитка
                            string t13 = textBox13.Text; //Температура входа в ЗВО
                            string t14 = textBox14.Text; //Температура выхода из ЗВО

                            float Bsl = Convert.ToSingle(l11);
                            float Csl = Convert.ToSingle(l12);
                            float Tmet = Convert.ToSingle(t13);
                            float Tvux = Convert.ToSingle(t14);

                            if ((Bsl <= 1.9) && (Bsl >= 1))
                            {
                                if ((Csl <= 0.25) && (Csl >= 0.15))
                                {
                                    if ((Tmet <= 1550) && (Tmet >= 1500))
                                    {
                                        if ((Tvux <= 900) && (Tvux >= 600))
                                        {
                                            this.chart1.Series[0].Points.Clear();
                                            this.chart1.Series[1].Points.Clear();
                                            this.chart1.Series[2].Points.Clear();

                                            k = 1;
                                            n = 1;
                                            p = 0;
                                            timer = 0;
                                            mintimer = 0;

                                            timer1.Enabled = true;
                                            timer1.Start();
                                            timer2.Enabled = true;
                                            timer2.Start();

                                        }
                                        else
                                        {
                                            MessageBox.Show("Min/max температура металла на выходе из ЗВО 600 - 900 °С", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Min/max температура металла 1500 - 1550 °С", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Min/max толщина 0,15 - 0,25 м", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Min/max ширина 1 - 1,9 м", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }

                        }
                        else
                        {
                            MessageBox.Show("Min/max длина 0,1 - 10 м", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Символ ',' не допустимо", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Символ '0' не допустимо", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Заповніть всі поля", "EMPTY", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            //---------------------------------------------------------------------------
            
        }

        //---------------------------------------------------------------------------
        private void timer1_Tick(object sender, EventArgs e)
        {
            float V, L, Alpha1, Alpha2, Kz, Ekr, GAll1, t1, t2, F;

            var l11 = textBox11.Text; //Ширина слитка
            var l12 = textBox12.Text; //Высота слитка
            var t13 = textBox13.Text; //Температура входа в ЗВО
            var t14 = textBox14.Text; //Температура выхода из ЗВО

            Bsl = Convert.ToSingle(l11);
            Csl = Convert.ToSingle(l12);
            Tmet = Convert.ToSingle(t13) + 200;
            Tvux = Convert.ToSingle(t14);

            V = 0;
            Tvux = Tvux + 100;
            Kz = 0.0025F;
            V = (1.1F + trackBar1.Value / 100.0F) / 60;
            L = 0;
            Ekr = Bsl * 0.3F;
            GAll1 = 0;

            var FF = V.ToString();       
            this.Text = FF;

            string z1 = textBox1.Text; //Зона 1
            string z2 = textBox2.Text; //Зона 2
            string z3 = textBox3.Text; //Зона 3
            string z4 = textBox4.Text; //Зона 4
            string z5 = textBox5.Text; //Зона 5
            string z6 = textBox6.Text; //Зона 6
            string z7 = textBox7.Text; //Зона 7
            string z8 = textBox8.Text; //Зона 8
            string z9 = textBox9.Text; //Зона 9
            string l8 = textBox10.Text; //Длина кристализатора
           
            // float l_8 = Convert.ToSingle(l8);

            Csl = Convert.ToSingle(l12);
            //ZVO[0].Length = Convert.ToDouble(l8);//Длина кристализатора
            ZVO[1].Length = Convert.ToSingle(z1);//Зона 1
            ZVO[2].Length = Convert.ToSingle(z2);//Зона 2
            ZVO[3].Length = Convert.ToSingle(z3);//Зона 3
            ZVO[4].Length = Convert.ToSingle(z4);//Зона 4
            ZVO[5].Length = Convert.ToSingle(z5);//Зона 5
            ZVO[6].Length = Convert.ToSingle(z6);//Зона 6
            ZVO[7].Length = Convert.ToSingle(z7);//Зона 7
            ZVO[8].Length = Convert.ToSingle(z8);//Зона 8
            ZVO[9].Length = Convert.ToSingle(z9);//Зона 9
            ZVO[10].Length = 0.8F;

            Lzvo = ZVO[1].Length + ZVO[2].Length + ZVO[3].Length + ZVO[4].Length + ZVO[5].Length + ZVO[6].Length + ZVO[7].Length + ZVO[8].Length + ZVO[9].Length;

            for (int i = 1; i <= 10; i++) //for (int i = 1; i <= 10; i++)
            {
                if (i == 1)
                {
                    ZVO[i].Tau1 = 0;
                }
                else
                    ZVO[i].Tau1 = ZVO[i - 1].Tau2;

                Tz[10] = (ZVO[10].Length) / (V * 60);
                Tz[1] = ((ZVO[1].Length) / (V * 60)) + Tz[10];
                Tz[2] = ((ZVO[2].Length) / (V * 60)) + Tz[1];
                Tz[3] = ((ZVO[3].Length) / (V * 60)) + Tz[2];
                Tz[4] = ((ZVO[4].Length) / (V * 60)) + Tz[3];
                Tz[5] = ((ZVO[5].Length) / (V * 60)) + Tz[4];
                Tz[6] = ((ZVO[6].Length) / (V * 60)) + Tz[5];
                Tz[7] = ((ZVO[7].Length) / (V * 60)) + Tz[6];
                Tz[8] = ((ZVO[8].Length) / (V * 60)) + Tz[7];
                Tz[9] = ((ZVO[9].Length) / (V * 60)) + Tz[8];

                Tz1[10] = (ZVO[10].Length) / (1.2F);
                Tz1[1] = ((ZVO[1].Length) / (1.2F)) + Tz1[10];
                Tz1[2] = ((ZVO[2].Length) / (1.2F)) + Tz1[1];
                Tz1[3] = ((ZVO[3].Length) / (1.2F)) + Tz1[2];
                Tz1[4] = ((ZVO[4].Length) / (1.2F)) + Tz1[3];
                Tz1[5] = ((ZVO[5].Length) / (1.2F)) + Tz1[4];
                Tz1[6] = ((ZVO[6].Length) / (1.2F)) + Tz1[5];
                Tz1[7] = ((ZVO[7].Length) / (1.2F)) + Tz1[6];
                Tz1[8] = ((ZVO[8].Length) / (1.2F)) + Tz1[7];
                Tz1[9] = ((ZVO[9].Length) / (1.2F)) + Tz1[8];

                E[i] = Csl * 100 * Convert.ToSingle(Math.Sqrt(Tz[i]));
                EZ[i] = Csl * 100 * Convert.ToSingle(Math.Sqrt(Tz1[i]));

                if (i == 1)
                    tpov[i] = Tmet - (170 + 190 * (0.8F / (V * 60)));
                else
                    tpov[i] = tpov[i - 1] - (tpov[i - 1] - Tvux) * Convert.ToSingle(Math.Pow((ZVO[i].Length * 0.5F / Lzvo), 0.2F));

                L = L + ZVO[i].Length;
                ZVO[i].Tau2 = L / V;
                Alpha1 = Alpha0(ZVO[i].Tau1);
                Alpha2 = Alpha0(ZVO[i].Tau2);
                ZVO[i].Alpha = (Alpha1 + Alpha2) / 2;
                t1 = ZVO[i].Tau1;
                t2 = ZVO[i].Tau2;
                ZVO[i].Etha = Ekr + 2 * Kz * (Convert.ToSingle(Math.Sqrt(t2 * t2 * t2)) - Convert.ToSingle(Math.Sqrt(t1 * t1 * t1))) / (3 * (t2 - t1));
                ZVO[i].Phi = (Bsl - 2 * ZVO[i].Etha) / Bsl;
                F = (2 * Bsl + 2 * Csl) * ZVO[i].Length;
                ZVO[i].For1 = F * ZVO[i].Phi;
                ZVO[i].gF = (ZVO[i].Alpha - 140) / 37;
                ZVO[i].G = (ZVO[i].gF * ZVO[i].For1) / 2.64F;
                ZVO[i].DeltaG = (ZVO[i].G - ZVO[i].GCurr) * 0.1F;
                Koeff[i] = 100 / ZVO[i].G;
                T[i] = Tz[i] * 60;

                this.chart1.Series[2].Points.AddXY(1100, 0);
            }

            float[] ee = new float[11]; //[10]
            ee[10] = (EZ[10] / T[10]) + 0.0172F;
            ee[1] = (EZ[10] - EZ[1]) / (T[10] - T[1]);
            for (int i = 2; i <= 9; i++)
                ee[i] = (EZ[i - 1] - EZ[i]) / (T[i - 1] - T[i]);

            Y[1] = 0;
            for (int i = 2; i <= T[10]; i++) Y[i] = Y[i - 1] + ee[10];
            for (int i = Convert.ToInt32(T[10]); i <= T[1]; i++) Y[i] = Y[i - 1] + ee[1];
            for (int i = Convert.ToInt32(T[1]); i <= T[2]; i++) Y[i] = Y[i - 1] + ee[2];
            for (int i = Convert.ToInt32(T[2]); i <= T[3]; i++) Y[i] = Y[i - 1] + ee[3];
            for (int i = Convert.ToInt32(T[3]); i <= T[4]; i++) Y[i] = Y[i - 1] + ee[4];
            for (int i = Convert.ToInt32(T[4]); i <= T[5]; i++) Y[i] = Y[i - 1] + ee[5];
            for (int i = Convert.ToInt32(T[5]); i <= T[6]; i++) Y[i] = Y[i - 1] + ee[6];
            for (int i = Convert.ToInt32(T[6]); i <= T[7]; i++) Y[i] = Y[i - 1] + ee[7];
            for (int i = Convert.ToInt32(T[7]); i <= T[8]; i++) Y[i] = Y[i - 1] + ee[8];
            for (int i = Convert.ToInt32(T[8]); i <= T[9]; i++) Y[i] = Y[i - 1] + ee[9];
            for (int i = Convert.ToInt32(T[9]); i <= T[9]; i++) Y[i] = Y[i - 1];

            ZVO[10].G = 900 * 3.14F * 0.0004F * 7 * 36 * V * 60;
            ZVO[10].DeltaG = (ZVO[10].G - ZVO[10].GCurr) * 0.1F;

            float GAll, TemPP;

            GAll = 0;

            float[] Temp = new float[11];
            float[] TemP = new float[10];

            float[] Qvn = new float[10];
            float[] Qizl = new float[10];
            float[] Qkonv = new float[10];
            float[] gop = new float[10];
            float[] Fop = new float[10];
            float[] Gvod = new float[10];
            float[] Gvoz = new float[10];

            for (int i = 1; i <= 10; i++)
            {
                while (Convert.ToSingle(Math.Abs(ZVO[i].GCurr - ZVO[i].G)) > 0.0001F)
                {
                    ZVO[i].GCurr = ZVO[i].GCurr + ZVO[i].DeltaG;

                    break;
                }
                Temp[i] = Koeff[i] * ZVO[i].GCurr;
            }

            TemP[1] = tpov[2] - Temp[1];
            label40.Text = Math.Round(TemP[1], 2).ToString() + " °С"; //= TemP[1] + " С";
            TemP[2] = tpov[3] - Temp[2];
            label50.Text = Math.Round(TemP[2], 2).ToString() + " °С";
            TemP[3] = tpov[4] - Temp[3];
            label51.Text = Math.Round(TemP[3], 2).ToString() + " °С";
            TemP[4] = tpov[5] - Temp[4];
            label52.Text = Math.Round(TemP[4], 2).ToString() + " °С";
            TemP[5] = tpov[6] - Temp[5];
            label53.Text = Math.Round(TemP[5], 2).ToString() + " °С";
            TemP[6] = tpov[7] - Temp[6];
            label54.Text = Math.Round(TemP[6], 2).ToString() + " °С";
            TemP[7] = tpov[8] - Temp[7];
            label55.Text = Math.Round(TemP[7], 2).ToString() + " °С";
            TemP[8] = tpov[9] - Temp[8];
            label56.Text = Math.Round(TemP[8], 2).ToString() + " °С";
            TemP[9] = tpov[10] - Temp[9];
            label57.Text = Math.Round(TemP[9], 2).ToString() + " °С";

            TemPP = tpov[1] - Temp[9];

            label86.Text = Math.Round(TemPP, 2).ToString() + " °С"; //"Температура: " + Math.Round(TemPP, 2).ToString() + " °С";       // TemPP + " С";
            label85.Text = Math.Round(ZVO[10].GCurr, 2).ToString() + " м3/ч"; //"Расход: " + Math.Round(ZVO[10].GCurr, 2).ToString() + " м3/ч"; //ZVO[10].GCurr + " м3/ч";

            Qvn[1] = 27 * ((Tmet - 100 - TemP[1]) / (0.001F * E[1]));
            Qizl[1] = 0.75F * 5.67F * (Convert.ToSingle(Math.Pow(((TemP[1] + 273) / 100), 4)) - 78.9F);
            Qkonv[1] = 6.16F * (TemP[1] - 25);
            gop[1] = (Qvn[1] - Qizl[1] - Qkonv[1]) / 50000;
            Fop[1] = 0.001F * 2 * ZVO[1].Length * ((Bsl * 1000) - (2 * E[1]));
            Gvod[1] = gop[1] * Fop[1];

            for (int i = 2; i <= 9; i++)
            {
                Qvn[i] = 27 * ((Tmet - 100 - TemP[i]) / (0.001F * E[i]));
                Qizl[i] = 0.75F * 5.67F * (Convert.ToSingle(Math.Pow(((TemP[i] + 273) / 100), 4)) - 78.9F);
                Qkonv[i] = 22.88F * (TemP[i] - 25);
                gop[i] = (Qvn[i] - Qizl[i] - Qkonv[i]) / 58000;
                Fop[i] = 0.001F * 2 * ZVO[i].Length * ((Bsl * 1000) - (2 * E[i]));
                Gvod[i] = gop[i] * Fop[i];
                Gvoz[i] = Gvod[i] * 8;
                GAll = GAll + Gvoz[i];
            }

            for (int i = 1; i <= 9; i++)
            {
                GAll1 = GAll1 + Gvod[i];
            }

            label79.Text = Math.Round(GAll1, 2).ToString() + " м3/ч";   //GAll1 + " м3/ч";

            label41.Text = Math.Round(Gvod[1], 2).ToString() + " м3/ч";
            label42.Text = Math.Round(Gvod[2], 2).ToString() + " м3/ч";
            label43.Text = Math.Round(Gvod[3], 2).ToString() + " м3/ч";
            label44.Text = Math.Round(Gvod[4], 2).ToString() + " м3/ч";
            label45.Text = Math.Round(Gvod[5], 2).ToString() + " м3/ч";
            label46.Text = Math.Round(Gvod[6], 2).ToString() + " м3/ч";
            label47.Text = Math.Round(Gvod[7], 2).ToString() + " м3/ч";
            label48.Text = Math.Round(Gvod[8], 2).ToString() + " м3/ч";
            label49.Text = Math.Round(Gvod[9], 2).ToString() + " м3/ч";

            label80.Text = Math.Round(GAll, 2).ToString() + " м3/ч"; 

            label59.Text = Math.Round(Gvoz[2], 2).ToString() + " м3/ч"; //Gvoz[2] + " м3/ч"; 
            label60.Text = Math.Round(Gvoz[3], 2).ToString() + " м3/ч";
            label61.Text = Math.Round(Gvoz[4], 2).ToString() + " м3/ч";
            label62.Text = Math.Round(Gvoz[5], 2).ToString() + " м3/ч";
            label63.Text = Math.Round(Gvoz[6], 2).ToString() + " м3/ч";
            label64.Text = Math.Round(Gvoz[7], 2).ToString() + " м3/ч";
            label65.Text = Math.Round(Gvoz[8], 2).ToString() + " м3/ч";
            label66.Text = Math.Round(Gvoz[9], 2).ToString() + " м3/ч";
        }
        //---------------------------------------------------------------------------
       
        //---------------------------------------------------------------------------
        private void timer2_Tick_1(object sender, EventArgs e)
        {
            {
                int g;
                g = 210 - hScrollBar1.Value;     //hScrollBar1->Position;
                timer2.Interval = g; //timer2->Interval = g; (int) g;

                if (k <= T[9]) //if (k <= T[9]) { Series1->AddXY(k, Csl * 500); k++; }
                {
                    this.chart1.Series[0].Points.AddXY(k, Csl * 500);
                    k++;
                }
                    
                if (n <= T[9]) // if (n <= T[9]) { Series2->AddXY(n, Y[n]); n++; }
                {
                    this.chart1.Series[1].Points.AddXY(n, Y[n]);
                    n++; 
                }            

                var VV = g.ToString();
                this.Text = VV;

                if (p <= T[9])
                {
                    timer++;
                    if (timer == 60)
                    {
                        mintimer++;
                        timer = 0;
                    }
                    label83.Text = mintimer + " мин";
                    label89.Text = timer + " сек";
                    p++;
                }
            }
        }
        //---------------------------------------------------------------------------

        private void MNLZ_Load(object sender, EventArgs e)
        {
            //timer3.Interval = 5000;
            timer3.Enabled = true;
            // timer3.Start();
            textBox16.BackColor = Color.Red;
        }

        //---------------------------------------------------------------------------
        private void timer3_Tick(object sender, EventArgs e)
        {
            //label1.Text = DateTime.Now.ToString("yyyy.MM.dd, HH.mm.ss");

            label92.Text = DateTime.Now.ToString("HH:mm:ss");
           // label94.Text = DateTime.Now.ToString("yyyy.MM.dd");
            label94.Text = DateTime.Now.ToString("dd.MM.yyyy");

            textBox15.Text = DateTime.Now.ToString("HH:mm:ss");
            textBox15.TextAlign = HorizontalAlignment.Center; //выравнивание текста по центру

        }

        //---------------------------------------------------------------------------

        //---------------------------------------------------------------------------
        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            float V1;
            V1 = 1.1F + trackBar1.Value / 100.0F;
            label84.Text = Math.Round(V1, 2).ToString() + " м/мин";
            this.chart1.Series[2].Points.Clear(); //Series3->Clear();
            this.chart1.Series[2].Points.AddXY(T[9] + 100, 0F); // Series3->AddXY(T[9] + 100, 0);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox2.Focus();
                return;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox3.Focus();
                return;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox4.Focus();
                return;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox5.Focus();
                return;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox6.Focus();
                return;
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox7.Focus();
                return;
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox10.Focus();
                return;
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox8.Focus();
                return;
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    textBox9.Focus();
                return;
            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
            }

            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                    button2.Focus();
                return;
            }
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Stream mystr = null;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((mystr = openFileDialog1.OpenFile()) != null)
                {
                    StreamReader myread = new StreamReader(mystr);
                    string[] str;
                    int num = 0;
                    try
                    {
                        string[] str1 = myread.ReadToEnd().Split('\n');
                        num = str1.Count();
                        dataGridView1.RowCount = num;
                        for (int i = 0; i < num; i++)
                        {
                            str = str1[i].Split(' ');
                            for (int j = 0; j < dataGridView1.ColumnCount; j++)
                            {
                                try
                                {
                                    string data = str[j].Replace("[etot_siavol]", " ");
                                    dataGridView1.Rows[i].Cells[j].Value = data;
                                    //dataGridView1.Rows[i].Cells[j].Value = str[j];
                                }
                                catch { }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        myread.Close();
                    }
                }
            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Stream myStream;
            //saveFileDialog1.Filter = "(*.doc) | *.doc|(*.docx) | *.docx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((myStream = saveFileDialog1.OpenFile()) != null)
                {
                    StreamWriter myWritet = new StreamWriter(myStream);
                    try
                    {
                        for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                        {
                            for (int j = 0; j < dataGridView1.ColumnCount; j++)
                            {
                                string data = dataGridView1.Rows[i].Cells[j].Value.ToString().Replace(" ", "[etot_siavol]");
                                myWritet.Write(data + " ");
                                //myWritet.Write(dataGridView1.Rows[i].Cells[j].Value.ToString() + " ");
                                //if((dataGridView1.ColumnCount - j) != 1) myWritet.Write("^");
                            }
                            myWritet.WriteLine();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        myWritet.Close();
                    }
                    myStream.Close();
                }
            }
        }
                
        private void exelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

      /*
        private void экспортВMSWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "(*.doc) | *.doc|(*.docx) | *.docx";
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            string filename = saveFileDialog1.FileName;


            var application = new Microsoft.Office.Interop.Word.Application();
            var document = new Microsoft.Office.Interop.Word.Document();
            document = application.Documents.Add();
            //application.Visible = true;
            Word.Paragraph p1;
            p1 = document.Content.Paragraphs.Add();
            p1.Range.Text = textBox1.Text + "\n";
            /*p1.Range.Text = "График функции sin(x)";
            p1.Range.InsertAfter("\n1 строка");
            p1.Range.InsertAfter("\n2 строка");
            p1.Range.InsertAfter("\n3 строка\n");*/
       /*
            Word.InlineShape pictureShape1 = p1.Range.InlineShapes.AddPicture(Directory.GetCurrentDirectory() + "//chart.png");
            Word.InlineShape pictureShape2 = p1.Range.InlineShapes.AddPicture(Directory.GetCurrentDirectory() + "//DataGridView.png");
            //MessageBox.Show(Directory.GetCurrentDirectory() + "\\chart.png");
            document.SaveAs(filename);
            //document.SaveAs("d:\\file.docx");
            application.Quit();
            application = null;
        }
       */
    }
}
