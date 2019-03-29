using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace AlwaysOnTopper
{
    static class Program
    {
        const Int32 MenuId = 31337;
        static string MenuItemName = "Always on top";
        const int Interval = 500;
        const int OffsetFromBottom = 1;

        static Dictionary<string, string> Languages = new Dictionary<string, string>()
        {
            {"ru", "Поверх всех окон" }
        };

        #region Unmanaged
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool GetMenuItemInfo(IntPtr hMenu, UInt32 uItem, bool fByPosition, [In, Out] MENUITEMINFO lpmii);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InsertMenuItem(IntPtr hMenu, uint uItem, bool fByPosition, [In] MENUITEMINFO lpmii);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool SetMenuItemInfo(IntPtr hMenu, uint uItem, bool fByPosition, [In] MENUITEMINFO lpmii);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool RemoveMenu(IntPtr hMenu, uint uItem, bool fByPosition);

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetMenuItemCount(IntPtr hMenu);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr
           hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess,
           uint idThread, uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetWindowInfo(IntPtr hwnd, ref WINDOWINFO pwi);

        [StructLayout(LayoutKind.Sequential)]
        struct WINDOWINFO
        {
            public uint cbSize;
            public RECT rcWindow;
            public RECT rcClient;
            public uint dwStyle;
            public uint dwExStyle;
            public uint dwWindowStatus;
            public uint cxWindowBorders;
            public uint cyWindowBorders;
            public ushort atomWindowType;
            public ushort wCreatorVersion;

            public WINDOWINFO(Boolean? filler) : this()   // Allows automatic initialization of "cbSize" with "new WINDOWINFO(null/true/false)".
            {
                cbSize = (UInt32)(Marshal.SizeOf(typeof(WINDOWINFO)));
            }

        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, int hWndInsertAfter,
            int X, int Y, int cx, int cy, SetWindowPosFlags uFlags);

        [Flags]
        private enum SetWindowPosFlags : uint
        {
            /// <summary>If the calling thread and the thread that owns the window are attached to different input queues, 
            /// the system posts the request to the thread that owns the window. This prevents the calling thread from 
            /// blocking its execution while other threads process the request.</summary>
            /// <remarks>SWP_ASYNCWINDOWPOS</remarks>
            AsynchronousWindowPosition = 0x4000,
            /// <summary>Prevents generation of the WM_SYNCPAINT message.</summary>
            /// <remarks>SWP_DEFERERASE</remarks>
            DeferErase = 0x2000,
            /// <summary>Draws a frame (defined in the window's class description) around the window.</summary>
            /// <remarks>SWP_DRAWFRAME</remarks>
            DrawFrame = 0x0020,
            /// <summary>Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to 
            /// the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE 
            /// is sent only when the window's size is being changed.</summary>
            /// <remarks>SWP_FRAMECHANGED</remarks>
            FrameChanged = 0x0020,
            /// <summary>Hides the window.</summary>
            /// <remarks>SWP_HIDEWINDOW</remarks>
            HideWindow = 0x0080,
            /// <summary>Does not activate the window. If this flag is not set, the window is activated and moved to the 
            /// top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter 
            /// parameter).</summary>
            /// <remarks>SWP_NOACTIVATE</remarks>
            DoNotActivate = 0x0010,
            /// <summary>Discards the entire contents of the client area. If this flag is not specified, the valid 
            /// contents of the client area are saved and copied back into the client area after the window is sized or 
            /// repositioned.</summary>
            /// <remarks>SWP_NOCOPYBITS</remarks>
            DoNotCopyBits = 0x0100,
            /// <summary>Retains the current position (ignores X and Y parameters).</summary>
            /// <remarks>SWP_NOMOVE</remarks>
            IgnoreMove = 0x0002,
            /// <summary>Does not change the owner window's position in the Z order.</summary>
            /// <remarks>SWP_NOOWNERZORDER</remarks>
            DoNotChangeOwnerZOrder = 0x0200,
            /// <summary>Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to 
            /// the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent 
            /// window uncovered as a result of the window being moved. When this flag is set, the application must 
            /// explicitly invalidate or redraw any parts of the window and parent window that need redrawing.</summary>
            /// <remarks>SWP_NOREDRAW</remarks>
            DoNotRedraw = 0x0008,
            /// <summary>Same as the SWP_NOOWNERZORDER flag.</summary>
            /// <remarks>SWP_NOREPOSITION</remarks>
            DoNotReposition = 0x0200,
            /// <summary>Prevents the window from receiving the WM_WINDOWPOSCHANGING message.</summary>
            /// <remarks>SWP_NOSENDCHANGING</remarks>
            DoNotSendChangingEvent = 0x0400,
            /// <summary>Retains the current size (ignores the cx and cy parameters).</summary>
            /// <remarks>SWP_NOSIZE</remarks>
            IgnoreResize = 0x0001,
            /// <summary>Retains the current Z order (ignores the hWndInsertAfter parameter).</summary>
            /// <remarks>SWP_NOZORDER</remarks>
            IgnoreZOrder = 0x0004,
            /// <summary>Displays the window.</summary>
            /// <remarks>SWP_SHOWWINDOW</remarks>
            ShowWindow = 0x0040,
        }

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [Flags]
        enum MIIM
        {
            BITMAP = 0x00000080,
            CHECKMARKS = 0x00000008,
            DATA = 0x00000020,
            FTYPE = 0x00000100,
            ID = 0x00000002,
            STATE = 0x00000001,
            STRING = 0x00000040,
            SUBMENU = 0x00000004,
            TYPE = 0x00000010
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        class MENUITEMINFO
        {
            public Int32 cbSize = Marshal.SizeOf(typeof(MENUITEMINFO));
            public MIIM fMask;
            public UInt32 fType;
            public UInt32 fState;
            public UInt32 wID;
            public IntPtr hSubMenu;
            public IntPtr hbmpChecked;
            public IntPtr hbmpUnchecked;
            public IntPtr dwItemData;
            public string dwTypeData = null;
            public UInt32 cch; // length of dwTypeData
            public IntPtr hbmpItem;

            public MENUITEMINFO() { }
            public MENUITEMINFO(MIIM pfMask)
            {
                fMask = pfMask;
            }
        }

        const uint EVENT_OBJECT_INVOKED = 0x8013;
        const uint WINEVENT_OUTOFCONTEXT = 0;
        const UInt32 MFT_STRING = 0x00000000;
        const UInt32 MFS_CHECKED = 0x00000008;
        const UInt32 MFS_UNCHECKED = 0x00000000;
        const int HWND_NOTOPMOST = -2;
        const int HWND_TOPMOST = -1;
        const int WS_EX_TOPMOST = 0x00000008;
        #endregion

        static Dictionary<IntPtr, IntPtr> windowHandles = new Dictionary<IntPtr, IntPtr>();

        static bool IsTopmost(IntPtr hwnd)
        {
            var info = new WINDOWINFO(true);
            GetWindowInfo(hwnd, ref info);
            return (info.dwExStyle & WS_EX_TOPMOST) != 0;
        }

        static void UpdateAlwaysOnTopToMenu(IntPtr windowHwnd, bool remove = false)
        {
            IntPtr sysMenu;
            int count;
            sysMenu = GetSystemMenu(windowHwnd, false);
            if ((count = GetMenuItemCount(sysMenu)) < 0) // Check if menu already modified
            {
                sysMenu = GetSystemMenu(windowHwnd, true);
                if ((count = GetMenuItemCount(sysMenu)) < 0)
                {
                    sysMenu = GetSystemMenu(windowHwnd, false);
                    if ((count = GetMenuItemCount(sysMenu)) < 0)
                        return;
                }
            }

            // Calculate target position
            uint position = (uint)Math.Max(0, count - OffsetFromBottom);

            // Check if it's already our menu item
            var item = new MENUITEMINFO(MIIM.STATE | MIIM.FTYPE | MIIM.ID | MIIM.STRING);
            item.dwTypeData = new string(' ', 64);
            item.cch = (uint)item.dwTypeData.Length;
            if (!GetMenuItemInfo(sysMenu, (uint)Math.Max(0, (int)position - 1), true, item))
                return;
            // Need to add new menu item?
            var newItem = item.dwTypeData != MenuItemName && item.wID != MenuId;
            var state = IsTopmost(windowHwnd) ? MFS_CHECKED : MFS_UNCHECKED;
            // Need to update menu item?
            var updateItem = !newItem && (
                    (state & (MFS_CHECKED | MFS_UNCHECKED))
                    != (item.fState & (MFS_CHECKED | MFS_UNCHECKED))
                );

            if (remove)
            {
                if (!newItem) // If menu item exists
                    //RemoveMenu(sysMenu, (uint)Math.Max(0, (int)position - 1), true);
                    GetSystemMenu(windowHwnd, true); // Reset menu
            }
            else if (newItem || updateItem)
            {
                item = new MENUITEMINFO(MIIM.STATE | MIIM.FTYPE | MIIM.ID | MIIM.STRING);
                item.fType = MFT_STRING;
                item.dwTypeData = MenuItemName;
                item.cch = (uint)item.dwTypeData.Length;
                item.fState = state;
                item.wID = MenuId;
                if (newItem) // Add menu item
                    InsertMenuItem(sysMenu, position, true, item);
                else if (updateItem) // Update menu item
                    SetMenuItemInfo(sysMenu, (uint)Math.Max(0, (int)position - 1), true, item);
            }

            if (remove) // Deattach hook?
            {
                var hooks = windowHandles.Where(kv => kv.Value == windowHwnd).Select(kv => kv.Key).ToArray();
                foreach (var hook in hooks)
                {
                    UnhookWinEvent(hook);
                    windowHandles.Remove(hook);
                }
            }
            // Attach hook to target window
            else if (!windowHandles.ContainsValue(windowHwnd))
            {
                var hhook = SetWinEventHook(EVENT_OBJECT_INVOKED, EVENT_OBJECT_INVOKED, windowHwnd,
                        WinEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);
                if (hhook != IntPtr.Zero)
                    windowHandles[hhook] = windowHwnd;
            }
        }

        static void WinEventProc(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (idChild != MenuId)
                return;
            IntPtr windowHwnd;
            if (!windowHandles.TryGetValue(hWinEventHook, out windowHwnd))
                return;
            if (GetForegroundWindow() != windowHwnd)
                return;
            SetWindowPos(windowHwnd, IsTopmost(windowHwnd) ? HWND_NOTOPMOST : HWND_TOPMOST,
                0, 0, 0, 0, SetWindowPosFlags.IgnoreMove | SetWindowPosFlags.IgnoreResize);
            UpdateAlwaysOnTopToMenu(windowHwnd); // Update menu
        }

        [STAThread]
        static void Main()
        {
            bool createdNew = true;
            var mutex = new Mutex(true, Assembly.GetExecutingAssembly().GetName().Name, out createdNew);
            if (!createdNew)
                return;
            try
            {
                CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
                Languages.TryGetValue(currentCulture.TwoLetterISOLanguageName, out MenuItemName);

                // Dumb form for message queue
                var form = new Form();
                form.Load += new EventHandler((object o, EventArgs e) =>
                {
                    form.WindowState = FormWindowState.Minimized;
                    form.ShowInTaskbar = false;
                });

                var timer = new System.Threading.Timer(new TimerCallback(
                    (object state) =>
                    {
                        try
                        {
                            form.Invoke(new Action(() =>
                            {
                                UpdateAlwaysOnTopToMenu(GetForegroundWindow());
                            }));
                        }
                        catch { }
                    }),
                null, Interval, Interval);
                Application.Run(form);
            }
            finally
            {
                mutex.ReleaseMutex();
                foreach (var hwnd in windowHandles.Values.ToArray())
                    UpdateAlwaysOnTopToMenu(hwnd, remove: true);
            }
        }
    }
}
