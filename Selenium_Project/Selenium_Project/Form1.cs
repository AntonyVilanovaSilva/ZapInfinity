namespace Selenium_Project
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           Selenium_Bot bot = new Selenium_Bot();
            bot.EnviarMensagem("Leonardo");
        }
    }
}