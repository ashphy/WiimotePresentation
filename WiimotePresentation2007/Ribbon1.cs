using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using WiimoteLib;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace WiimotePresentation2007
{
    public partial class Ribbon1
    {
        WiimoteCollection mWC;

        /// <summary>
        /// In now a slide show mode?
        /// </summary>
        bool isSlideShowMode = false;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //Search wiimote at start up
            wiiRemoteInit();
        }

        /// <summary>
        /// Initializing a wiimote
        /// </summary>
        void wiiRemoteInit()
        {
            disconnectWiimote();

            message.Label = "Connecting...";
            mWC = new WiimoteCollection();

            try
            {
                //Search wiimote
                mWC.FindAllWiimotes();

                //Connect the wiimote
                int index = 1;
                foreach (Wiimote wm in mWC)
                {
                    wm.Disconnect();
                    wm.WiimoteChanged += new EventHandler<WiimoteChangedEventArgs>(wm_WiimoteChanged);
                    wm.Connect();
                    wm.SetLEDs(index++);
                    refreshBattery(wm.WiimoteState);
                }

                connectWiimoteButton.Checked = true;
            }
            catch (WiimoteNotFoundException)
            {
                //MessageBox.Show(ex.Message, "Wiimote not found error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                message.Label = "Wiimote not found";
                connectWiimoteButton.Checked = false;
                return;
            }
            catch (WiimoteException)
            {
                //MessageBox.Show(ex.Message, "Wiimote error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                message.Label = "Wiimote not found";
                connectWiimoteButton.Checked = false;
                return;
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message, "Unknown error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                message.Label = "Unknown error";
                connectWiimoteButton.Checked = false;
                return;
            }

            Globals.ThisAddIn.Application.SlideShowBegin += new EApplication_SlideShowBeginEventHandler(Application_SlideShowBegin);
            Globals.ThisAddIn.Application.SlideShowEnd += new EApplication_SlideShowEndEventHandler(Application_SlideShowEnd);
        }

        private void disconnectWiimote()
        {
            if (mWC != null)
            {
                foreach (Wiimote wm in mWC)
                {
                    wm.Disconnect();
                    wm.Dispose();
                }
            }

            message.Label = "Disconected wiimote";
        }

        /// <summary>
        /// End of the slide show
        /// </summary>
        /// <param name="Pres"></param>
        void Application_SlideShowEnd(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            isSlideShowMode = false;
            timer.Enabled = false;
        }

        /// <summary>
        /// Start of the slide show
        /// </summary>
        /// <param name="Wn"></param>
        void Application_SlideShowBegin(Microsoft.Office.Interop.PowerPoint.SlideShowWindow Wn)
        {
            isSlideShowMode = true;

            //Start the slideshow timer
            if (presentationTimerEnable.Checked)
            {
                timer.Enabled = true;
                timer.Interval = int.Parse(timerInterval.Text) * 1000;
            }
        }

        void wm_WiimoteChanged(object sender, WiimoteChangedEventArgs e)
        {
            //Does thw wiimote use?
            if (!connectWiimoteButton.Checked) return;
            if (Globals.ThisAddIn == null) return;

            //Wiimote state
            WiimoteState ws = e.WiimoteState;

            Presentation ap = Globals.ThisAddIn.Application.ActivePresentation;

            try
            {

                if (isSlideShowMode)
                {
                    SlideShowView view = ap.SlideShowWindow.View;

                    if (ws.ButtonState.One && ws.ButtonState.Two)
                    {
                        //Exit slide show
                        view.Exit();
                    }
                    else if (ws.ButtonState.Left)
                    {
                        //Back to the slide show
                        view.Previous();
                    }
                    else if (ws.ButtonState.A)
                    {
                        //Go to the next slide
                        view.Next();
                    }
                    else if (ws.ButtonState.Home)
                    {
                        //Go to slide sorter view
                        view.Exit();
                        DocumentWindow aw = Globals.ThisAddIn.Application.ActiveWindow;
                        aw.ViewType = PpViewType.ppViewSlideSorter;
                    }
                }
                else if (Globals.ThisAddIn.Application.ActiveWindow.ViewType == PpViewType.ppViewSlideSorter)
                {
                    DocumentWindow aw = Globals.ThisAddIn.Application.ActiveWindow;

                    //If slide sorter view
                    if (ws.ButtonState.Up)
                    {
                        SendKeys.SendWait("{UP}");
                    }
                    else if (ws.ButtonState.Down)
                    {
                        SendKeys.SendWait("{down}");
                    }
                    else if (ws.ButtonState.Left)
                    {
                        SendKeys.SendWait("{LEFT}");
                    }
                    else if (ws.ButtonState.Right)
                    {
                        SendKeys.SendWait("{RIGHT}");
                    }
                    else if (ws.ButtonState.Home)
                    {
                        aw.ViewType = PpViewType.ppViewNormal;
                    }
                    else if (ws.ButtonState.A)
                    {
                        showPresentationFromSelection();
                    }
                    else if (ws.ButtonState.Minus)
                    {
                        //Scalling down
                        int current = aw.View.Zoom;
                        if (0 < current - 10)
                        {
                            aw.View.Zoom = current - 10;
                        }
                    }
                    else if (ws.ButtonState.Plus)
                    {
                        //Scalling up
                        int current = aw.View.Zoom;

                        if (current + 10 <= 100)
                        {
                            aw.View.Zoom = current + 10;
                        }
                        else if (current + 10 > 100)
                        {
                            aw.View.Zoom = 100;
                        }
                    }
                }
                else
                {
                    DocumentWindow aw = Globals.ThisAddIn.Application.ActiveWindow;
                    int currentPage = aw.Selection.SlideRange.SlideIndex;

                    //Normal view mode
                    if (ws.ButtonState.Up)
                    {
                        if (1 < currentPage)
                        {
                            ap.Slides.Range(currentPage - 1).Select();
                        }
                    }
                    else if (ws.ButtonState.Down)
                    {
                        if (currentPage < ap.Slides.Count)
                        {
                            ap.Slides.Range(currentPage + 1).Select();
                        }
                    }
                    else if (ws.ButtonState.A)
                    {
                        showPresentationFromSelection();
                    }
                    else if (ws.ButtonState.Home)
                    {
                        aw.ViewType = PpViewType.ppViewSlideSorter;
                    }
                }
            }
            catch (COMException ce)
            {
                Console.WriteLine(ce.StackTrace);
            }

            refreshBattery(ws);
        }
                
        /// <summary>
        /// Start slide show from the selection slide
        /// </summary>
        private void showPresentationFromSelection()
        {
            //int currentSlide = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
            //Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.StartingSlide = 2;
            //battery.Label = Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.StartingSlide.ToString();
            //Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.Run();
            SendKeys.SendWait("+{F5}");
        }

        /// <summary>
        /// Update the battery value
        /// </summary>
        private void refreshBattery(WiimoteState ws)
        {
            message.Label = "Battery:" + ws.Battery.ToString("F1") + "%";
        }

        private void connectWiimoteButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (connectWiimoteButton.Checked)
            {
                wiiRemoteInit();
            }
            else
            {
                disconnectWiimote();
            }
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            //Presentation timer
            foreach (Wiimote wm in mWC)
            {
                //vibrator
                wm.SetRumble(true);
                System.Threading.Thread.Sleep(1000);
                wm.SetRumble(false);
            }
        }
    }
}
