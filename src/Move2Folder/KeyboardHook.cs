using System;
using System.Runtime.InteropServices;
using System.Windows.Threading;

namespace Move2Folder
{
    public sealed class KeyboardHook : IDisposable
    {
        private const int WH_KEYBOARD = 2;
        private const int VK_SHIFT = 0x10;
        private const int VK_CONTROL = 0x11;
        private const int VK_MENU = 0x12; // Alt
        private const int VK_V = 0x56;

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        private readonly HookProc _hookProc;
        private IntPtr _hookHandle = IntPtr.Zero;
        private readonly Action _onHotkeyPressed;
        private readonly Dispatcher _dispatcher;
        private bool _disposed;

        public KeyboardHook(Action onHotkeyPressed)
        {
            _onHotkeyPressed = onHotkeyPressed ?? throw new ArgumentNullException(nameof(onHotkeyPressed));
            _dispatcher = Dispatcher.CurrentDispatcher;

            // Keep reference to delegate to avoid garbage collection
            _hookProc = KeyboardHookCallback;

            uint threadId = GetCurrentThreadId();
            _hookHandle = SetWindowsHookEx(WH_KEYBOARD, _hookProc, IntPtr.Zero, threadId);

            if (_hookHandle == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                System.Diagnostics.Debug.WriteLine($"Failed to set keyboard hook. Error: {error}");
            }
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                long l = lParam.ToInt64();
                // Bit 31 indicates transition state: 0 = key down, 1 = key up
                bool isKeyDown = (l & (1L << 31)) == 0;
                // Bit 30 indicates previous state
                bool isRepeat = (l & (1L << 30)) != 0;

                int vkCode = wParam.ToInt32();

                if (isKeyDown && !isRepeat && vkCode == VK_V)
                {
                    bool isShiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                    bool isAltPressed = (GetKeyState(VK_MENU) & 0x8000) != 0;
                    bool isCtrlPressed = (GetKeyState(VK_CONTROL) & 0x8000) != 0;

                    bool isCtrlShiftV = isCtrlPressed && isShiftPressed && !isAltPressed;

                    if (isCtrlShiftV)
                    {
                        // Trigger hotkey action asynchronously on dispatcher
                        _dispatcher.BeginInvoke(new Action(() =>
                        {
                            try
                            {
                                _onHotkeyPressed();
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error invoking hotkey action: {ex}");
                            }
                        }), DispatcherPriority.Input);

                        // Suppress standard processing of Ctrl+Shift+V (overriding Outlook default)
                        return (IntPtr)1;
                    }
                }
            }

            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                if (_hookHandle != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hookHandle);
                    _hookHandle = IntPtr.Zero;
                }
            }
        }
    }
}
