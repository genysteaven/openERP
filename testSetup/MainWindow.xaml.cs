using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.IO.Compression;
using System.Net;


namespace testSetup
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void bt_start_Click(object sender, RoutedEventArgs e)
        {
            //installation de VirtualBox 
            if (InstallVBox())
            {
                MessageBox.Show("Installation de VirtualBox terminée. Lancement de l'importation de la VM OpenERP.");
            }
            else 
            {
                MessageBox.Show("Erreur lors de l'installation de virtualBox...");
            }

            //importation de la VM OpenERP
            //et installation des extensions
            if (ImportVM())
            {
                MessageBox.Show("Importation de la VM OpenERP terminée. Lancement de l'installation des extensions");
            }
            else
            {
                MessageBox.Show("Erreur lors de l'installation de virtuaBox...");
            }

            //installation des guestAdditions
            if (InstallGuestAdd())
            {
                MessageBox.Show("Installation des outils complémentaires terminée.");
            }
            else
            {
                MessageBox.Show("Erreur lors de l'installation de virtuaBox...");
            }

            //installation de PgAdmin
            if(InstallPgAdmin())
            {
                MessageBox.Show("Installation de PgAdmin terminée.");
            }
            else
            {
                MessageBox.Show("Erreur lors de l'installation de PgAdmin...");
            }


        }

        private Boolean InstallVBox()
        {
            //chemin de l'exe de VirtualBox. Attention, il faut mettre des / et non des \ qui ne sont pas reconnus par C#
            string cheminVBox = Environment.GetEnvironmentVariable("HOMEPATH") + "/Desktop/ressourcesOpenERP/VirtualBox-4.3.0-89960-Win.exe";

            //repertoire d'execution
            string cheminExec = Environment.GetEnvironmentVariable("HOMEPATH") + "/Desktop/ressourcesOpenERP";
            //MessageBox.Show(cheminVBox);
            //MessageBox.Show(cheminExec);

            //lancement de l'installation SILENCIEUSE (argument /s) de VirtualBox
            //string cheminAppData = Environment.GetEnvironmentVariable("APPDATA");

            //extraction des fichiers dans C:\VBox (inutile de créer ce répertoire avant)
            //voir pour décommenter ou non
            //ExecuteExe(cheminVBox, cheminExec, "-extract -path " + cheminExec);

            //recupération de l'architecture du systeme hote (x86 ou amd64)
            String leMsi;
            bool is64bit = !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"));
            if (is64bit)
            {
                //MessageBox.Show("archi 64");
                leMsi = "VirtualBox-4.3.0-r89960-MultiArch_amd64.msi";
            }
            else
            {
               // MessageBox.Show("archi 32");
                leMsi = "VirtualBox-4.3.0-r89960-MultiArch_x86.msi";
            }

            //installation du certificat
            //verifier si c'est obligatoire
            String cheminCertUtil = Environment.GetEnvironmentVariable("WINDIR") + "/System32/certutil.exe";
            ExecuteExe(cheminCertUtil, cheminExec, " -addstore \"TrustedPublisher\" oracle-vbox.cert");

            //utiliser cette sans faire de extract et enlever la ligne utilisant msiexec ? A tester
            ExecuteExe(cheminVBox, cheminExec, "-s");

            //execution du programme msiexec
            //String cheminMsiExec = Environment.GetEnvironmentVariable("WINDIR") + "/System32/msiexec.exe";
            //ExecuteExe(cheminMsiExec, cheminExec, " /i " + leMsi + " APPLOCAL=VBoxApploication,VBoxUSB,VBoxNetwork /qb");
            //remarque : l'option n pour ne pas avoir d'interface utilisateur à l'installation ne fonctionne pas...

            //MessageBox.Show("terminé");
            return true;
        }

        //private Boolean InstallExtPack()
        //{
        //    string cheminRessources = Environment.GetEnvironmentVariable("HOMEPATH") + "/Desktop/ressourcesOpenERP";
        //    //installation du pack extension
        //    string cheminVBoxManage = Environment.GetEnvironmentVariable("PROGRAMFILES") + "/Oracle/VirtualBox/VBoxManage.exe";
        //    ExecuteExe(cheminVBoxManage, cheminRessources, "extpack install Oracle_VM_VirtualBox_Extension_Pack-4.3.0-89960.vbox-extpack");
           
        //    return true;
        //}

        private Boolean InstallGuestAdd()
        {
            string cheminRessources = Environment.GetEnvironmentVariable("HOMEPATH") + "/Desktop/ressourcesOpenERP";
            string zipPath = cheminRessources + "/VBoxGuestAdditions.iso";
            //Decompression de VBoxGuestAdditions.iso
            ZipFile.ExtractToDirectory(zipPath, cheminRessources + "/");
            //installation des guest additions
            //Le repertoire VBoxGuestAddition est obtenu en desarchivant le VBoxGuestAdditions.iso (automatiser cela ?? ...) 
            ExecuteExe(cheminRessources + "/VBoxGuestAdditions/VBoxWindowsAdditions.exe", cheminRessources + "/VBoxGuestAdditions", "/S");

            return true;
        }

        private Boolean ImportVM()
        {
            string cheminRessources = Environment.GetEnvironmentVariable("HOMEPATH") + "/Desktop/ressourcesOpenERP";
            //importation de la VM OpenERP
            //voir pour cacher invite de commande
            string cheminVBoxManage = Environment.GetEnvironmentVariable("PROGRAMFILES") + "/Oracle/VirtualBox/VBoxManage.exe";
            ExecuteExe(cheminVBoxManage, cheminRessources, " import export-OpenERP-Server.ova");

            //installation du pack extension
            ExecuteExe(cheminVBoxManage, cheminRessources, "extpack install Oracle_VM_VirtualBox_Extension_Pack-4.3.0-89960.vbox-extpack");
           
            return true;
        }

        private Boolean InstallPgAdmin()
        {
            string cheminPg = Environment.GetEnvironmentVariable("HOMEPATH") + "/Desktop/ressourcesOpenERP/pgadmin";
            ExecuteExe(cheminPg, cheminPg + "/pgadmin3", "");
           
            return true;
        }

        //méthode permettant l'execution d'un .exe
        private static void ExecuteExe(string process_, string workingDirectory_, string args_)
        {
            ProcessStartInfo psi = new ProcessStartInfo(process_)
            {
                UseShellExecute = true,
                RedirectStandardOutput = false,
                RedirectStandardInput = false,
                RedirectStandardError = false,
                CreateNoWindow = true,
                WorkingDirectory = workingDirectory_,
                Arguments = args_
            };
            Process leProcessus = Process.Start(psi);
            leProcessus.WaitForExit();
        }
    }
}
