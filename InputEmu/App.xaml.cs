using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace InputEmu
{
  /// <summary>
  /// Interaction logic for App.xaml
  /// </summary>
  public partial class App : Application
  {
    public static String[]? Args { get; set; }
    private void Application_Startup(Object sender, StartupEventArgs e)
    {
      Args = e.Args;
    }
  }
}
