using BeamLab.SPEF.Test.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BeamLab.SPEF.Test
{
    public partial class Form1 : Form
    {
        private SPEFContext repository;
        private SPEFContext Repository
        {
            get
            {
                if (repository == null)
                {
                    var mainVariationLabel = "en-gb";
                    repository = new TestContext(txtUrl.Text, mainVariationLabel);
                    repository.SPNetworkCredentials = new System.Net.NetworkCredential(txtUsername.Text, txtPassword.Text, txtDomain.Text);
                }
                return repository;
            }
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void btnInit_Click(object sender, EventArgs e)
        {
            var msg = Repository.Init();
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            var allNews = Repository.Load<News>();

            foreach(var news in allNews)
            {
                if(news.Category != null && news.Category.ID > 0)
                {
                    var category = Repository.LoadByID<Category>(news.Category.ID);
                }
            }
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            var category = new Category()
            {
                Title = $"Categoria {DateTime.Now.Minute}",
                Position = 1
            };

            var newItemID = Repository.Save(category);
            if (newItemID >= 0)
            {
                category = Repository.LoadByID<Category>(newItemID);


                var news = new News()
                {
                    Title = $"News della categoria {newItemID}",
                    Text = $"Testo della news della categoria {newItemID}",
                    Category = category
                };
                var newNewsID = Repository.Save(news);
            }

        }
    }
}
