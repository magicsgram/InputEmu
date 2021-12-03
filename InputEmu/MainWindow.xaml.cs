using Nefarius.ViGEm.Client;
using Nefarius.ViGEm.Client.Targets;
using Nefarius.ViGEm.Client.Targets.Xbox360;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Vortice.XInput;

namespace InputEmu
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    private readonly Int32 maxMessageQueueSize = 500;
    private readonly Queue<String> messageTextBoxQueue = new();
    private readonly IXbox360Controller virtualController;
    private Single inputDuration = 0.1f;
    private Single inputInterval = 0.5f;
    private Boolean abortScriptNow = false;
    private Boolean processImmersiveMode = false;
    private Int32 immersiveScriptCount = 0;
    private readonly List<List<(TimeSpan timeStamp, State state)>> inputLists = new();

    // dotnet publish -c Release -r win-x64 --self-contained -p:PublishReadyToRun=true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
    public MainWindow()
    {
      ViGEmClient client = new();
      virtualController = client.CreateXbox360Controller();
      virtualController.Connect();
      virtualController.AutoSubmitReport = false;

      this.Loaded += async (sender, e) =>
      {
        mainWindow.Title = "Initializing...";
        Int32 userIndex = 0;
        await Task.Run(() =>
        {
          for (; ; )
          {
            try
            {
              userIndex = virtualController.UserIndex;
              break;
            }
            catch { }
          }
          if (App.Args?.Length > 0)
          {
            LogToMessageBox("Initializating...");
            LogToMessageBox($"Loading {App.Args[0]}.");
            if (LoadInputFile(App.Args[0]))
              LogToMessageBox("Done.");
            else
              LogToMessageBox("Fail.");
          }
        });
        commandTextBox.IsEnabled = true;
        mainWindow.Title = $"InputEmu3 (Player {virtualController.UserIndex + 1})";
        LogToMessageBox($"Current Player is {virtualController.UserIndex + 1}");
        LogToMessageBox("Command Ready.");
      };
      this.Closing += (sender, e) =>
      {
        try
        {
          CleanUp(virtualController);
          virtualController.Disconnect();
          client.Dispose();
        }
        catch { }
      };
      InitializeComponent();
    }

    private void LogToMessageBox(String s)
    {
      Dispatcher.Invoke(() =>
      {
        messageTextBoxQueue.Enqueue(s);
        while (messageTextBoxQueue.Count >= maxMessageQueueSize)
          messageTextBoxQueue.Dequeue();

        StringBuilder stringBuilder = new();
        foreach (String s in messageTextBoxQueue.ToList())
          stringBuilder.AppendLine(s);

        messageTextBox.Text = stringBuilder.ToString();
        messageTextBox.ScrollToEnd();
      }, System.Windows.Threading.DispatcherPriority.Send);
    }

    private async void CommandTextBox_KeyDown(Object sender, KeyEventArgs e)
    {
      if (e.Key == Key.Enter)
      {
        TextBox textBox = (TextBox)sender;
        String commandInput = textBox.Text;
        if (commandInput.Length > 0)
        {
          LogToMessageBox("> " + commandInput);
          textBox.Text = "Executing.....";
          commandTextBox.IsEnabled = false;
          String[] commands = commandInput.Split(' ');

          WindowsHook.IKeyboardMouseEvents? globalHook = null;
          globalHook = WindowsHook.Hook.GlobalEvents();
          globalHook.KeyDown += GlobalHook_KeyDown;
          globalHook.KeyUp += GlobalHook_KeyUp;
          await Task<String>.Run(() => ProcessCommand(commands));
          globalHook.KeyDown -= GlobalHook_KeyDown;
          globalHook.KeyUp -= GlobalHook_KeyUp;
          globalHook.Dispose();

          virtualController.ResetReport();
          virtualController.SubmitReport();

          textBox.Text = "";
          commandTextBox.IsEnabled = true;
        }
      }
    }

    private void GlobalHook_KeyUp(Object sender, WindowsHook.KeyEventArgs e) => HandleKeyHook(e, false);

    private void GlobalHook_KeyDown(Object sender, WindowsHook.KeyEventArgs e) => HandleKeyHook(e, true);

    private void HandleKeyHook(WindowsHook.KeyEventArgs e, Boolean keyPressed)
    {
      switch (e.KeyCode)
      {
        case WindowsHook.Keys.Q:
          abortScriptNow = true;
          return;
        case WindowsHook.Keys.Oemtilde:
          processImmersiveMode = false;
          abortScriptNow = true;
          return;
      }

      if (!processImmersiveMode)
        return;

      Xbox360Button? button = null;
      Xbox360Slider? slider = null;
      Byte sliderValue = 0;
      Xbox360Axis? axis = null;
      Int16 axisValue = 0;
      switch (e.KeyCode)
      {
        case WindowsHook.Keys.R:
          immersiveScriptCount = Int32.MaxValue;
          return;
        case WindowsHook.Keys.D1:
          immersiveScriptCount = 1;
          return;
        case WindowsHook.Keys.OemQuestion:
          button = Xbox360Button.A;
          break;
        case WindowsHook.Keys.OemQuotes:
          button = Xbox360Button.B;
          break;
        case WindowsHook.Keys.OemPeriod:
          button = Xbox360Button.X;
          break;
        case WindowsHook.Keys.OemSemicolon:
          button = Xbox360Button.Y;
          break;
        case WindowsHook.Keys.Enter:
          button = Xbox360Button.Start;
          break;
        case WindowsHook.Keys.RShiftKey:
          button = Xbox360Button.Back;
          break;
        case WindowsHook.Keys.L:
          button = Xbox360Button.LeftShoulder;
          break;
        case WindowsHook.Keys.P:
          button = Xbox360Button.RightShoulder;
          break;
        case WindowsHook.Keys.Q:
          button = Xbox360Button.LeftThumb;
          break;
        case WindowsHook.Keys.E:
          button = Xbox360Button.RightThumb;
          break;
        case WindowsHook.Keys.Y:
          button = Xbox360Button.Guide;
          break;
        case WindowsHook.Keys.Oemcomma:
          slider = Xbox360Slider.LeftTrigger;
          sliderValue = 255;
          break;
        case WindowsHook.Keys.OemOpenBrackets:
          slider = Xbox360Slider.RightTrigger;
          sliderValue = 255;
          break;

        case WindowsHook.Keys.S:
          button = Xbox360Button.Up;
          break;
        case WindowsHook.Keys.X:
          button = Xbox360Button.Right;
          break;
        case WindowsHook.Keys.Z:
          button = Xbox360Button.Down;
          break;
        case WindowsHook.Keys.A:
          button = Xbox360Button.Left;
          break;
        case WindowsHook.Keys.G:
          button = Xbox360Button.LeftThumb;
          break;
        case WindowsHook.Keys.H:
          button = Xbox360Button.RightThumb;
          break;

        case WindowsHook.Keys.F:
          axis = Xbox360Axis.LeftThumbY;
          axisValue = 32767;
          break;
        case WindowsHook.Keys.V:
          axis = Xbox360Axis.LeftThumbX;
          axisValue = 32767;
          break;
        case WindowsHook.Keys.C:
          axis = Xbox360Axis.LeftThumbY;
          axisValue = -32768;
          break;
        case WindowsHook.Keys.D:
          axis = Xbox360Axis.LeftThumbX;
          axisValue = -32768;
          break;
        case WindowsHook.Keys.J:
          axis = Xbox360Axis.RightThumbY;
          axisValue = 32767;
          break;
        case WindowsHook.Keys.K:
          axis = Xbox360Axis.RightThumbX;
          axisValue = 32767;
          break;
        case WindowsHook.Keys.M:
          axis = Xbox360Axis.RightThumbY;
          axisValue = -32768;
          break;
        case WindowsHook.Keys.N:
          axis = Xbox360Axis.RightThumbX;
          axisValue = -32768;
          break;

      }
      if (button == null && slider == null && axis == null)
        return;

      if (button != null)
        virtualController.SetButtonState(button, keyPressed);
      else if (axis != null)
      {
        if (!keyPressed)
          axisValue = 0;
        virtualController.SetAxisValue(axis, axisValue);
      }
      else if (slider != null)
      {
        if (!keyPressed)
          sliderValue = 0;
        virtualController.SetSliderValue(slider, sliderValue);
      }

      virtualController.SubmitReport();
    }

    private void ProcessCommand(String[] commands)
    {
      abortScriptNow = false;
      if (commands[0] == "run")
      {
        LogToMessageBox("Waiting 2 seconds to run");
        Thread.Sleep(2000);
        Int32 upperBound = Int32.MaxValue;
        if (commands.Length > 1)
          upperBound = Int32.Parse(commands[1]);

        ExecuteLoadedScript(upperBound);
      }
      else if (commands[0] == "immersive")
      {
        LogToMessageBox("Starting immersive mode.");
        processImmersiveMode = true;
        for (; ; )
        {
          if (!processImmersiveMode)
            break;

          if (immersiveScriptCount > 0)
          {

            LogToMessageBox($"Running loaded script: {(immersiveScriptCount == Int32.MaxValue ? "infinite" : immersiveScriptCount)}.");
            ExecuteLoadedScript(immersiveScriptCount);
            LogToMessageBox("Done.");
            immersiveScriptCount = 0;
            abortScriptNow = false;
          }

          Thread.Sleep(250);
        }
        LogToMessageBox("Exiting immersive mode");
      }
      else if (commands[0] == "load")
      {
        LogToMessageBox($"Loading {commands[1]}.");
        if (LoadInputFile(commands[1]))
          LogToMessageBox("Done.");
        else
          LogToMessageBox("Fail.");
      }
      else if (commands[0] == "duration")
      {
        if (!Single.TryParse(commands[1], out Single durationParsed))
          LogToMessageBox($"Duration cannot be parsed.");
        else
        {
          LogToMessageBox($"Duration is now {durationParsed} seconds.");
          inputDuration = durationParsed;
        }
      }
      else if (commands[0] == "interval")
      {
        if (!Single.TryParse(commands[1], out Single intervalParsed))
          LogToMessageBox($"Interval cannot be parsed.");
        else
        {
          LogToMessageBox($"Interval is now {intervalParsed} seconds.");
          inputInterval = intervalParsed;
        }
      }
      else if (commands[0] == "clean")
      {
        LogToMessageBox("Cleaning.");
        CleanUp(virtualController);
        LogToMessageBox("Done.");
      }
      else if (commands[0] == "keep")
      {
        TimeSpan delay = TimeSpan.FromSeconds(5);
        TimeSpan duration = TimeSpan.FromSeconds(0.3);
        while (!abortScriptNow)
        {
          Thread.Sleep(delay);
          virtualController.ResetReport();
          virtualController.SubmitReport();
          virtualController.SetButtonState(Xbox360Button.LeftThumb, true);
          virtualController.SubmitReport();
          Thread.Sleep(duration);
          virtualController.SetButtonState(Xbox360Button.LeftThumb, false);
          virtualController.SubmitReport();
          Thread.Sleep(duration);
          virtualController.SetButtonState(Xbox360Button.RightThumb, true);
          virtualController.SubmitReport();
          Thread.Sleep(duration);
          virtualController.SetButtonState(Xbox360Button.RightThumb, false);
          virtualController.SubmitReport();
        }
      }
      else
      {
        TimeSpan delay = TimeSpan.FromSeconds(2);
        TimeSpan duration = TimeSpan.FromSeconds(inputDuration);
        Thread.Sleep(delay);
        LogToMessageBox("Executing.");

        for (Int32 i = 0; i < commands.Length; ++i)
        {
          Xbox360Button? button = null;
          switch (commands[i])
          {
            case "a":
              button = Xbox360Button.A;
              break;
            case "b":
              button = Xbox360Button.B;
              break;
            case "x":
              button = Xbox360Button.X;
              break;
            case "y":
              button = Xbox360Button.Y;
              break;
            case "u":
              button = Xbox360Button.Up;
              break;
            case "r":
              button = Xbox360Button.Right;
              break;
            case "d":
              button = Xbox360Button.Down;
              break;
            case "l":
              button = Xbox360Button.Left;
              break;
            case "m":
              button = Xbox360Button.Start;
              break;
            case "v":
              button = Xbox360Button.Back;
              break;
            case "l1":
              button = Xbox360Button.LeftShoulder;
              break;
            case "r1":
              button = Xbox360Button.RightShoulder;
              break;
            case "l3":
              button = Xbox360Button.LeftThumb;
              break;
            case "r3":
              button = Xbox360Button.RightThumb;
              break;
            case "g":
              button = Xbox360Button.Guide;
              break;
          }
          if (button == null)
            continue;

          LogToMessageBox($"Press Q to abort to quit. {i + 1} / {commands.Length} : {button.Name}");
          virtualController.ResetReport();
          virtualController.SetButtonState(button, true);
          virtualController.SubmitReport();
          Thread.Sleep(duration);
          virtualController.SetButtonState(button, false);
          virtualController.SubmitReport();

          if (abortScriptNow)
          {
            LogToMessageBox("Aborting.");
            return;
          }

          if (i < commands.Length - 1)
            Thread.Sleep(TimeSpan.FromSeconds(inputInterval));
        }
      }

      virtualController.ResetReport();
      virtualController.SubmitReport(); ;
    }

    private void ExecuteLoadedScript(Int32 upperBound)
    {
      Stopwatch totalTimeStopWatch = new();
      totalTimeStopWatch.Start();
      TimeSpan oneSecondTimeSpan = TimeSpan.FromSeconds(1d);
      for (Int32 runCount = 1; runCount <= upperBound; ++runCount)
      {
        for (Int32 i = 0; runCount <= upperBound && i < inputLists.Count; ++i)
        {
          List<(TimeSpan timeStamp, State state)> inputList = inputLists[i];
          LogToMessageBox($"Press Q to abort. {totalTimeStopWatch.Elapsed}. Run {runCount} : {i + 1} / {inputLists.Count}");
          Stopwatch inputTimer = new();
          inputTimer.Start();
          for (Int32 j = 1; runCount <= upperBound && j < inputList.Count; ++j)
          {
            State state = inputList[j].state;
            virtualController.ResetReport();
            virtualController.SetButtonState(Xbox360Button.A, state.Gamepad.Buttons.HasFlag(GamepadButtons.A));
            virtualController.SetButtonState(Xbox360Button.B, state.Gamepad.Buttons.HasFlag(GamepadButtons.B));
            virtualController.SetButtonState(Xbox360Button.X, state.Gamepad.Buttons.HasFlag(GamepadButtons.X));
            virtualController.SetButtonState(Xbox360Button.Y, state.Gamepad.Buttons.HasFlag(GamepadButtons.Y));
            virtualController.SetButtonState(Xbox360Button.LeftShoulder, state.Gamepad.Buttons.HasFlag(GamepadButtons.LeftShoulder));
            virtualController.SetButtonState(Xbox360Button.RightShoulder, state.Gamepad.Buttons.HasFlag(GamepadButtons.RightShoulder));
            virtualController.SetButtonState(Xbox360Button.LeftThumb, state.Gamepad.Buttons.HasFlag(GamepadButtons.LeftThumb));
            virtualController.SetButtonState(Xbox360Button.RightThumb, state.Gamepad.Buttons.HasFlag(GamepadButtons.RightThumb));
            virtualController.SetButtonState(Xbox360Button.Left, state.Gamepad.Buttons.HasFlag(GamepadButtons.DPadLeft));
            virtualController.SetButtonState(Xbox360Button.Right, state.Gamepad.Buttons.HasFlag(GamepadButtons.DPadRight));
            virtualController.SetButtonState(Xbox360Button.Up, state.Gamepad.Buttons.HasFlag(GamepadButtons.DPadUp));
            virtualController.SetButtonState(Xbox360Button.Down, state.Gamepad.Buttons.HasFlag(GamepadButtons.DPadDown));
            virtualController.SetButtonState(Xbox360Button.Back, state.Gamepad.Buttons.HasFlag(GamepadButtons.Back));
            virtualController.SetButtonState(Xbox360Button.Start, state.Gamepad.Buttons.HasFlag(GamepadButtons.Start));

            virtualController.SetAxisValue(Xbox360Axis.LeftThumbX, state.Gamepad.LeftThumbX);
            virtualController.SetAxisValue(Xbox360Axis.LeftThumbY, state.Gamepad.LeftThumbY);
            virtualController.SetSliderValue(Xbox360Slider.LeftTrigger, state.Gamepad.LeftTrigger);
            virtualController.SetAxisValue(Xbox360Axis.RightThumbX, state.Gamepad.RightThumbX);
            virtualController.SetAxisValue(Xbox360Axis.RightThumbY, state.Gamepad.RightThumbY);
            virtualController.SetSliderValue(Xbox360Slider.RightTrigger, state.Gamepad.RightTrigger);

            for (; ; )
            {
              if (abortScriptNow)
              {
                LogToMessageBox("Aborting.");
                return;
              }

              TimeSpan timeDiff = inputList[j].timeStamp - inputTimer.Elapsed;
              if (timeDiff < TimeSpan.Zero)
                break;
              if (timeDiff > oneSecondTimeSpan)
                Thread.Sleep(500);
              else
              {
                Thread.Sleep(timeDiff);
                break;
              }
            }

            virtualController.SubmitReport();
          }
        }
      }
    }

    private Boolean LoadInputFile(String filePath)
    {
      try
      {
        inputLists.Clear();
        StreamReader sr = new(filePath);
        while (!sr.EndOfStream)
        {
#pragma warning disable CS8604 // Possible null reference argument.
          FileStream fs = new(sr.ReadLine(), FileMode.Open);
#pragma warning restore CS8604 // Possible null reference argument.
          BinaryReader br = new(fs);
          List<(TimeSpan, State)> inputList = new();
          while (br.BaseStream.Position < br.BaseStream.Length)
          {
            TimeSpan timespan = TimeSpan.FromTicks(br.ReadInt64());
            int packetNumber = br.ReadInt32();

            GamepadButtons flags = (GamepadButtons)br.ReadInt32();
            short leftThumbX = br.ReadInt16();
            short leftThumbY = br.ReadInt16();
            byte leftTrigger = br.ReadByte();
            short rightThumbX = br.ReadInt16();
            short rightThumbY = br.ReadInt16();
            byte rightTrigger = br.ReadByte();

            Gamepad newGamepad = new()
            {
              Buttons = flags,
              LeftThumbX = leftThumbX,
              LeftThumbY = leftThumbY,
              LeftTrigger = leftTrigger,
              RightThumbX = rightThumbX,
              RightThumbY = rightThumbY,
              RightTrigger = rightTrigger
            };
            State newState = new()
            {
              PacketNumber = packetNumber,
              Gamepad = newGamepad
            };

            inputList.Add((timespan, newState));
          }

          br.Close();
          fs.Close();
          inputLists.Add(inputList);
        }
        sr.Close();
        return true;
      }
      catch { return false; }
    }

    private static void CleanUp(IXbox360Controller virtualController)
    {
      Single duration = 0.08f;

      virtualController.ResetReport();
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetAxisValue(Xbox360Axis.LeftThumbX, (Int16)32767);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetAxisValue(Xbox360Axis.LeftThumbX, (Int16)0);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetAxisValue(Xbox360Axis.LeftThumbY, (Int16)32767);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetAxisValue(Xbox360Axis.LeftThumbY, (Int16)0);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetAxisValue(Xbox360Axis.RightThumbX, (Int16)32767);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetAxisValue(Xbox360Axis.RightThumbX, (Int16)0);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetAxisValue(Xbox360Axis.RightThumbY, (Int16)32767);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetAxisValue(Xbox360Axis.RightThumbY, (Int16)0);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetSliderValue(Xbox360Slider.LeftTrigger, 255);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetSliderValue(Xbox360Slider.LeftTrigger, 0);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetSliderValue(Xbox360Slider.RightTrigger, 255);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetSliderValue(Xbox360Slider.RightTrigger, 0);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.A, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.A, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.B, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.B, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.X, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.X, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.Y, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.Y, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.Up, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.Up, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.Down, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.Down, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.Left, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.Left, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.Right, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.Right, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.RightShoulder, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.RightShoulder, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.LeftShoulder, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.LeftShoulder, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.RightThumb, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.RightThumb, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.LeftThumb, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.LeftThumb, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.Start, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.Start, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      virtualController.SetButtonState(Xbox360Button.Back, true);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));
      virtualController.SetButtonState(Xbox360Button.Back, false);
      virtualController.SubmitReport();
      Thread.Sleep(TimeSpan.FromSeconds(duration));

      //virtualController.SetButtonState(Xbox360Button.Guide, true);
      //virtualController.SubmitReport();
      //Thread.Sleep(TimeSpan.FromSeconds(duration));
      //virtualController.SetButtonState(Xbox360Button.Guide, false);
      //virtualController.SubmitReport();
      //Thread.Sleep(TimeSpan.FromSeconds(duration));
    }
  }
}
