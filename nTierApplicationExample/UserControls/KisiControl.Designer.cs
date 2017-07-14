namespace nTierApplicationExample.UserControls
{
    partial class KisiControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose (bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent ()
        {
            this.id = new System.Windows.Forms.TextBox();
            this.ad = new System.Windows.Forms.TextBox();
            this.soyad = new System.Windows.Forms.TextBox();
            this.yas = new System.Windows.Forms.TextBox();
            this.sehir = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.adres = new System.Windows.Forms.TextBox();
            this.ilce = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // id
            // 
            this.id.Location = new System.Drawing.Point(3, 29);
            this.id.Name = "id";
            this.id.Size = new System.Drawing.Size(100, 20);
            this.id.TabIndex = 0;
            // 
            // ad
            // 
            this.ad.Location = new System.Drawing.Point(3, 74);
            this.ad.Name = "ad";
            this.ad.Size = new System.Drawing.Size(100, 20);
            this.ad.TabIndex = 1;
            // 
            // soyad
            // 
            this.soyad.Location = new System.Drawing.Point(3, 119);
            this.soyad.Name = "soyad";
            this.soyad.Size = new System.Drawing.Size(100, 20);
            this.soyad.TabIndex = 2;
            // 
            // yas
            // 
            this.yas.Location = new System.Drawing.Point(3, 164);
            this.yas.Name = "yas";
            this.yas.Size = new System.Drawing.Size(100, 20);
            this.yas.TabIndex = 3;
            // 
            // sehir
            // 
            this.sehir.Location = new System.Drawing.Point(3, 254);
            this.sehir.Name = "sehir";
            this.sehir.Size = new System.Drawing.Size(100, 20);
            this.sehir.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(18, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "ID";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(22, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Adı";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 100);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(39, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Soyadı";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 145);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(25, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Yaş";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 235);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(31, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Şehir";
            // 
            // adres
            // 
            this.adres.Location = new System.Drawing.Point(3, 209);
            this.adres.Name = "adres";
            this.adres.Size = new System.Drawing.Size(100, 20);
            this.adres.TabIndex = 10;
            // 
            // ilce
            // 
            this.ilce.Location = new System.Drawing.Point(3, 299);
            this.ilce.Name = "ilce";
            this.ilce.Size = new System.Drawing.Size(100, 20);
            this.ilce.TabIndex = 11;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(3, 190);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(34, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "Adres";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(3, 280);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(24, 13);
            this.label7.TabIndex = 13;
            this.label7.Text = "İlçe";
            // 
            // KisiControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.ilce);
            this.Controls.Add(this.adres);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.sehir);
            this.Controls.Add(this.yas);
            this.Controls.Add(this.soyad);
            this.Controls.Add(this.ad);
            this.Controls.Add(this.id);
            this.Name = "KisiControl";
            this.Size = new System.Drawing.Size(110, 331);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.TextBox id;
        public System.Windows.Forms.TextBox ad;
        public System.Windows.Forms.TextBox soyad;
        public System.Windows.Forms.TextBox yas;
        public System.Windows.Forms.TextBox sehir;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.Label label4;
        public System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        public System.Windows.Forms.TextBox adres;
        public System.Windows.Forms.TextBox ilce;

    }
}
