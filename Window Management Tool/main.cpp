// Include necessary headers
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <vector>
#include <string>
#include <map>
#include <cmath>
#include <float.h>

// Link necessary libraries
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "Comctl32.lib")

// Constants for control positioning
#define MARGIN 10
#define BUTTON_HEIGHT 30
#define BUTTON_WIDTH 150
#define SMALL_BUTTON_WIDTH 100
#define EDIT_WIDTH 50
#define LABEL_HEIGHT 20
#define LABEL_WIDTH 200
#define COMBOBOX_WIDTH 200
#define COMBOBOX_HEIGHT 100
#define MOVE_BUTTON_WIDTH 100
#define MOVE_BUTTON_HEIGHT 30
#define MIN_WINDOW_WIDTH 700
#define MIN_WINDOW_HEIGHT 500

// Control IDs
#define ID_CAPTURE_BUTTON 1
#define ID_RESTORE_BUTTON 2
#define ID_CLEAR_BUTTON 4
#define ID_ARRANGE_BUTTON 5
#define ID_NUM_WINDOWS_EDIT 3
#define ID_MONITOR_COMBOBOX 6
#define ID_WINDOW_LISTVIEW 7
#define ID_MOVE_UP_BUTTON 8
#define ID_MOVE_DOWN_BUTTON 9
#define ID_PIXEL_FIX_X_LABEL 10    // Control ID for Horizontal Pixel Fix Label
#define ID_PIXEL_FIX_X_EDIT 11     // Control ID for Horizontal Pixel Fix Edit Box
#define ID_PIXEL_FIX_Y_LABEL 12    // Control ID for Vertical Pixel Fix Label
#define ID_PIXEL_FIX_Y_EDIT 13     // Control ID for Vertical Pixel Fix Edit Box
#define ID_MIN_SPACING_Y_LABEL 14  // Control ID for Min Vertical Spacing Label
#define ID_MIN_SPACING_Y_EDIT 15   // Control ID for Min Vertical Spacing Edit Box

// *** New Control IDs ***
#define ID_WINDOWTITLE_LABEL 18          // Label for Window Title
#define ID_WINDOWTITLE_EDIT 16           // Edit box for Window Title input
#define ID_CAPTURE_BY_TITLE_BUTTON 17    // Button to capture windows by title

// Custom message for unhooking
#define WM_UNHOOK_HOOKS (WM_USER + 1)

// Structure to store window information
struct WindowInfo
{
    HWND hWnd;
    RECT rect;
    std::wstring windowTitle;
    int index; // Index number in the order of registration
};

// Global variables
std::vector<WindowInfo> windowList;
bool isCapturing = false;
bool captureCompleted = false;
int windowsToCapture = 0;
int windowsCaptured = 0;
HHOOK hMouseHook = NULL;
HHOOK hKeyboardHook = NULL;
HHOOK hGlobalKeyboardHook = NULL; // For global shortcuts
HINSTANCE hInstance;
HWND hListView;
HWND hMonitorComboBox;
HWND hMainWindow;

// Control handles
HWND hCaptureButton;
HWND hRestoreButton;
HWND hClearButton;
HWND hArrangeButton;
HWND hNumWindowsEdit;
HWND hNumWindowsLabel;
HWND hMonitorLabel;
HWND hMoveUpButton;
HWND hMoveDownButton;
HWND hPixelFixXLabel;   // Handle for Horizontal Pixel Fix Label
HWND hPixelFixXEdit;    // Handle for Horizontal Pixel Fix Edit Box
HWND hPixelFixYLabel;   // Handle for Vertical Pixel Fix Label
HWND hPixelFixYEdit;    // Handle for Vertical Pixel Fix Edit Box
HWND hMinSpacingYLabel; // Handle for Min Vertical Spacing Label
HWND hMinSpacingYEdit;  // Handle for Min Vertical Spacing Edit Box

// *** New Control Handles ***
HWND hWindowTitleLabel;          // Label for Window Title
HWND hWindowTitleEdit;           // Edit box for Window Title input
HWND hCaptureByTitleButton;      // Button to capture windows by title

// Original window procedure for ListView
WNDPROC OldListViewProc = NULL;

// Global variables for message box timeout
UINT_PTR g_msgboxTimerId = 0;
HHOOK g_hMsgBoxHook = NULL; // Corrected variable name
HWND g_hMsgBoxWnd = NULL;   // Handle to the message box window

// Map to store monitor information
std::map<int, MONITORINFO> monitorMap;

// Function prototypes
LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);
BOOL CALLBACK MonitorEnumProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM dwData);
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam); // New: Callback for EnumWindows
void UpdateMonitorComboBox();
void RefreshWindowList();
void MoveSelectedItem(int direction);
LRESULT CALLBACK ListViewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
void AdjustControls();
void ArrangeWindows();
VOID CALLBACK MsgBoxTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
LRESULT CALLBACK CBTProc(int nCode, WPARAM wParam, LPARAM lParam);
int TimedMessageBox(HWND hWnd, LPCTSTR lpText, LPCTSTR lpCaption, UINT uType);
void StartCapturingWindows(HWND hWnd);
void RestoreWindows();
void ClearCapturedWindows();
HWND GetWindowUnderCursor();
void CaptureWindowUnderCursor();
void CheckAndRemoveClosedWindows();
void CaptureWindowsByTitle(const std::wstring& title); // New: Function to capture windows by title

// Function to get the window under the cursor
HWND GetWindowUnderCursor()
{
    POINT pt;
    GetCursorPos(&pt);
    HWND hWnd = WindowFromPoint(pt);

    // If it's a child window, get the parent window
    while (hWnd && (GetWindowLong(hWnd, GWL_STYLE) & WS_CHILD))
    {
        hWnd = GetParent(hWnd);
    }
    return hWnd;
}

// Function to capture the window under the cursor
void CaptureWindowUnderCursor()
{
    HWND hWnd = GetWindowUnderCursor();
    if (hWnd && IsWindow(hWnd))
    {
        // Check if the window is already captured
        for (const auto& info : windowList)
        {
            if (info.hWnd == hWnd)
            {
                // Window already captured, no action needed
                return;
            }
        }

        WindowInfo info;
        info.hWnd = hWnd;
        GetWindowRect(hWnd, &info.rect);

        // Get window title
        wchar_t title[256];
        GetWindowText(hWnd, title, sizeof(title) / sizeof(wchar_t));
        info.windowTitle = title;

        // Set index number
        info.index = static_cast<int>(windowList.size()) + 1;

        windowList.push_back(info);
        windowsCaptured++;

        // Refresh the ListView
        RefreshWindowList();

        if (windowsCaptured >= windowsToCapture)
        {
            // Capturing complete
            captureCompleted = true;

            // Post a message to unhook the hooks after the hook procedure returns
            PostMessage(hMainWindow, WM_UNHOOK_HOOKS, 0, 0);
        }
    }
    else
    {
        // Invalid window, no action needed
    }
}

// Mouse hook callback
LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION)
    {
        if (wParam == WM_LBUTTONDOWN)
        {
            if (isCapturing)
            {
                // Capture window under the cursor
                CaptureWindowUnderCursor();

                // Suppress the click event
                return 1; // Stop processing this click
            }
            else if (captureCompleted)
            {
                // Suppress the click event
                return 1; // Stop processing this click
            }
        }
    }
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
}

// Keyboard hook callback
LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION && isCapturing)
    {
        KBDLLHOOKSTRUCT* pkbhs = (KBDLLHOOKSTRUCT*)lParam;
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN)
        {
            if (pkbhs->vkCode == VK_ESCAPE)
            {
                captureCompleted = true;

                // Post a message to unhook the hooks after the hook procedure returns
                PostMessage(hMainWindow, WM_UNHOOK_HOOKS, 0, 0);

                // Do not call MessageBox here

                return 1; // Prevent further processing
            }
        }
    }
    return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
}
LRESULT CALLBACK GlobalKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION)
    {
        KBDLLHOOKSTRUCT* pkbhs = (KBDLLHOOKSTRUCT*)lParam;
        if (wParam == WM_KEYDOWN && pkbhs->vkCode == 'J' && GetAsyncKeyState(VK_CONTROL) & 0x8000)
        {
            //OutputDebugString(L"Ctrl + J detected\n");
            ArrangeWindows();
            return 1; // Prevent further processing
        }
    }
    return CallNextHookEx(hGlobalKeyboardHook, nCode, wParam, lParam);
}

// Function to start capturing windows
void StartCapturingWindows(HWND hWnd)
{
    if (isCapturing)
    {
        TimedMessageBox(NULL, L"Window capturing is already in progress.", L"Info", MB_OK);
        return;
    }

    // Get the number of windows to capture
    wchar_t buffer[10];
    GetWindowText(hNumWindowsEdit, buffer, 10);
    windowsToCapture = _wtoi(buffer);
    if (windowsToCapture <= 0)
    {
        TimedMessageBox(NULL, L"Please enter a valid number of windows to capture.", L"Error", MB_OK);
        return;
    }

    // Initialize variables
    isCapturing = true;
    captureCompleted = false;
    windowsCaptured = 0;

    // Clear previous window list
    windowList.clear();
    ListView_DeleteAllItems(hListView);

    // Set mouse hook
    hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, LowLevelMouseProc, hInstance, 0);
    // Set keyboard hook
    hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, hInstance, 0);

    if (hMouseHook == NULL || hKeyboardHook == NULL)
    {
        TimedMessageBox(NULL, L"Cannot set hooks.", L"Error", MB_OK);
        isCapturing = false;
    }
    else
    {
        TimedMessageBox(NULL, L"Click on the windows you wish to capture. Press ESC to cancel.", L"Info", MB_OK);
    }
}

// *** New Function: Capture Windows by Title ***
void CaptureWindowsByTitle(const std::wstring& title)
{
    if (title.empty())
    {
        TimedMessageBox(NULL, L"Please enter a window title to capture.", L"Error", MB_OK);
        return;
    }

    // Clear previous window list
    windowList.clear();
    ListView_DeleteAllItems(hListView);

    // Enumerate all top-level windows and capture those with matching title
    EnumWindows(EnumWindowsProc, reinterpret_cast<LPARAM>(&title));

    // Refresh the ListView to display the updated windowList
    RefreshWindowList();

    if (windowList.empty())
    {
        TimedMessageBox(NULL, L"No windows found with the specified title.", L"Info", MB_OK);
    }
    else
    {
        TimedMessageBox(NULL, L"All matching windows have been captured.", L"Info", MB_OK);
    }
}

// *** New Callback Function: EnumWindowsProc ***
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam)
{
    std::wstring targetTitle = *(std::wstring*)lParam;

    // Get window title
    wchar_t windowTitle[256];
    GetWindowText(hwnd, windowTitle, sizeof(windowTitle) / sizeof(wchar_t));
    std::wstring currentTitle = windowTitle;

    // Check if the window title matches (case-insensitive)
    if (_wcsicmp(currentTitle.c_str(), targetTitle.c_str()) == 0)
    {
        // Check if the window is already captured
        for (const auto& info : windowList)
        {
            if (info.hWnd == hwnd)
            {
                // Already captured, skip
                return TRUE;
            }
        }

        // Add to windowList
        WindowInfo info;
        info.hWnd = hwnd;
        GetWindowRect(hwnd, &info.rect); // *** Corrected handle reference ***
        info.windowTitle = currentTitle;
        info.index = static_cast<int>(windowList.size()) + 1;
        windowList.push_back(info);
    }

    return TRUE; // Continue enumeration
}

// Function to restore window positions
void RestoreWindows()
{
    for (const auto& info : windowList)
    {
        if (IsWindow(info.hWnd))
        {
            SetWindowPos(info.hWnd, NULL, info.rect.left, info.rect.top,
                info.rect.right - info.rect.left,
                info.rect.bottom - info.rect.top,
                SWP_NOZORDER | SWP_NOACTIVATE);
        }
        else
        {
            std::wstring message = L"Window no longer available: " + info.windowTitle;
            TimedMessageBox(NULL, message.c_str(), L"Warning", MB_OK);
        }
    }
}

// Function to clear captured windows
void ClearCapturedWindows()
{
    windowList.clear();
    ListView_DeleteAllItems(hListView);
    TimedMessageBox(NULL, L"All captured windows have been cleared.", L"Info", MB_OK);
}

// Function to check and remove closed windows
void CheckAndRemoveClosedWindows()
{
    bool listChanged = false;
    for (auto it = windowList.begin(); it != windowList.end(); )
    {
        if (!IsWindow(it->hWnd))
        {
            std::wstring message = L"Window has been closed: " + it->windowTitle;
            TimedMessageBox(NULL, message.c_str(), L"Info", MB_OK);

            it = windowList.erase(it);
            listChanged = true;
        }
        else
        {
            ++it;
        }
    }
    if (listChanged)
    {
        // Update indices
        for (size_t i = 0; i < windowList.size(); ++i)
        {
            windowList[i].index = static_cast<int>(i) + 1;
        }
        // Refresh the ListView
        RefreshWindowList();
    }
}

// Function to arrange windows considering multiple monitors and ensuring equal sizes
void ArrangeWindows()
{
    if (windowList.empty())
    {
        TimedMessageBox(NULL, L"No windows to arrange. Please capture windows first.", L"Info", MB_OK);
        return;
    }

    // Check and remove any closed windows before arranging
    CheckAndRemoveClosedWindows();

    if (windowList.empty())
    {
        TimedMessageBox(NULL, L"No valid windows to arrange.", L"Info", MB_OK);
        return;
    }

    // Get selected monitor
    int monitorIndex = ComboBox_GetCurSel(hMonitorComboBox);
    if (monitorMap.find(monitorIndex) == monitorMap.end())
    {
        TimedMessageBox(NULL, L"Selected monitor not found.", L"Error", MB_OK);
        return;
    }
    MONITORINFO mi = monitorMap[monitorIndex];
    RECT workArea = mi.rcWork;
    int workWidth = workArea.right - workArea.left;
    int workHeight = workArea.bottom - workArea.top;

    // Retrieve pixelFix values
    wchar_t pixelFixXBuffer[16];
    GetWindowText(hPixelFixXEdit, pixelFixXBuffer, 16);
    int pixelFixX = _wtoi(pixelFixXBuffer); // Converts to integer, handles negative values

    wchar_t pixelFixYBuffer[16];
    GetWindowText(hPixelFixYEdit, pixelFixYBuffer, 16);
    int pixelFixY = _wtoi(pixelFixYBuffer); // Converts to integer, handles negative values

    // Retrieve minimum vertical spacing
    wchar_t minSpacingYBuffer[16];
    GetWindowText(hMinSpacingYEdit, minSpacingYBuffer, 16);
    int minSpacingY = _wtoi(minSpacingYBuffer);

    // Validate minimum spacing
    if (minSpacingY < 0)
    {
        TimedMessageBox(NULL, L"Minimum vertical spacing cannot be negative. Resetting to 0.", L"Invalid Input", MB_OK | MB_ICONWARNING);
        minSpacingY = 0;
    }

    // Apply pixelFixX to starting X position
    int startXPos = workArea.left + pixelFixX;

    // Apply pixelFixY to available height by reserving space at the bottom
    int totalSpacingY = (windowList.size() > 1) ? (windowList.size() - 1) * minSpacingY : 0;
    int availableHeight = workHeight - pixelFixY - totalSpacingY;

    // Ensure availableHeight is positive
    if (availableHeight <= 0)
    {
        TimedMessageBox(NULL, L"Not enough vertical space for the specified spacing and Pixel Fix Y. Please reduce the spacing or Pixel Fix Y.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Calculate the best grid layout (rows and columns) to minimize gaps
    int numWindows = static_cast<int>(windowList.size());
    double aspectRatio = static_cast<double>(workWidth) / (availableHeight);

    int bestRows = 1;
    int bestCols = numWindows;
    double minDifference = DBL_MAX;

    for (int rows = 1; rows <= numWindows; ++rows)
    {
        int cols = (numWindows + rows - 1) / rows; // Ceiling division
        double gridAspectRatio = static_cast<double>(cols) / rows;
        double difference = fabs(gridAspectRatio - aspectRatio);

        if (difference < minDifference)
        {
            minDifference = difference;
            bestRows = rows;
            bestCols = cols;
        }
    }

    // Recalculate total spacing based on bestRows
    totalSpacingY = (bestRows - 1) * minSpacingY;
    availableHeight = workHeight - pixelFixY - totalSpacingY;

    // Calculate window sizes
    int windowHeight = availableHeight / bestRows;
    int extraPixelsHeight = availableHeight % bestRows;

    // Distribute extra pixels to the top rows to eliminate gaps
    std::vector<int> rowHeights(bestRows, windowHeight);
    for (int row = 0; row < extraPixelsHeight; ++row)
    {
        rowHeights[row]++;
    }

    // Calculate window widths
    int windowWidth = workWidth / bestCols;
    int extraPixelsWidth = workWidth % bestCols;

    std::vector<int> columnWidths(bestCols, windowWidth);
    for (int col = 0; col < extraPixelsWidth; ++col)
    {
        columnWidths[col]++;
    }

    // Now, arrange the windows
    int windowIndex = 0;
    int yPos = workArea.top; // Start at the top of the work area
    for (int row = 0; row < bestRows; ++row)
    {
        int xPos = startXPos;
        for (int col = 0; col < bestCols; ++col)
        {
            if (windowIndex >= numWindows)
                break;

            const auto& info = windowList[windowIndex];

            if (IsWindow(info.hWnd))
            {
                int cellWidth = columnWidths[col];
                int cellHeight = rowHeights[row];

                // Desired client area size
                RECT desiredClientRect = { 0, 0, cellWidth, cellHeight };

                // Get current window style
                LONG style = GetWindowLong(info.hWnd, GWL_STYLE);
                LONG exStyle = GetWindowLong(info.hWnd, GWL_EXSTYLE);

                // Adjust window size to fit desired client area
                RECT adjustedWindowRect = desiredClientRect;
                if (!AdjustWindowRectEx(&adjustedWindowRect, style, FALSE, exStyle))
                {
                    // Fallback if AdjustWindowRectEx fails
                    adjustedWindowRect.right = cellWidth;
                    adjustedWindowRect.bottom = cellHeight;
                }

                int windowWidthAdjusted = adjustedWindowRect.right - adjustedWindowRect.left;
                int windowHeightAdjusted = adjustedWindowRect.bottom - adjustedWindowRect.top;

                // Move and resize the window
                BOOL success = SetWindowPos(
                    info.hWnd,
                    HWND_TOP,
                    xPos,
                    yPos,
                    windowWidthAdjusted,
                    windowHeightAdjusted,
                    SWP_NOACTIVATE | SWP_SHOWWINDOW
                );

                ShowWindow(info.hWnd, SW_MINIMIZE);
                // Minimize all windows first
                //for (const auto& info : windowList)
                //{
                //    if (IsWindow(info.hWnd))
                //    {
                //        ShowWindow(info.hWnd, SW_MINIMIZE);
                //    }
                //}
                Sleep(100);
                // Unminimize the window
                ShowWindow(info.hWnd, SW_RESTORE);
                ShowWindow(hMainWindow, SW_MINIMIZE);
                if (!success)
                {
                    // Optional: Handle the error, e.g., log or notify the user
                }

                // Bring the window to the foreground
                SetForegroundWindow(info.hWnd);
            }
            else
            {
                // Window is no longer valid; skip
                continue;
            }

            // Move to the next column position
            xPos += columnWidths[col];
            windowIndex++;
        }
        // After arranging a row, move Y position down by row height + spacing
        yPos += rowHeights[row] + minSpacingY;
    }
}

// Callback for enumerating monitors
BOOL CALLBACK MonitorEnumProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM dwData)
{
    MONITORINFO mi = { 0 };
    mi.cbSize = sizeof(mi);
    if (GetMonitorInfo(hMonitor, &mi))
    {
        int index = static_cast<int>(monitorMap.size());
        monitorMap[index] = mi;

        wchar_t monitorName[256];
        wsprintf(monitorName, L"Monitor %d (%dx%d)", index + 1,
            mi.rcMonitor.right - mi.rcMonitor.left,
            mi.rcMonitor.bottom - mi.rcMonitor.top);

        ComboBox_AddString((HWND)dwData, monitorName);
    }
    return TRUE;
}

// Function to update monitor combo box
void UpdateMonitorComboBox()
{
    ComboBox_ResetContent(hMonitorComboBox);
    monitorMap.clear();
    EnumDisplayMonitors(NULL, NULL, MonitorEnumProc, (LPARAM)hMonitorComboBox);
    ComboBox_SetCurSel(hMonitorComboBox, 0);
}

// Function to refresh the window list
void RefreshWindowList()
{
    ListView_DeleteAllItems(hListView);

    for (const auto& info : windowList)
    {
        // Format the index number
        wchar_t indexStr[16];
        swprintf_s(indexStr, 16, L"%d", info.index);

        // Format the window handle
        wchar_t handleStr[32];
        swprintf_s(handleStr, 32, L"0x%016IX", (UINT_PTR)info.hWnd);

        // Insert index column
        LVITEM lvItem = { 0 };
        lvItem.mask = LVIF_TEXT | LVIF_PARAM;
        lvItem.pszText = indexStr;
        lvItem.lParam = (LPARAM)info.hWnd;
        lvItem.iItem = ListView_GetItemCount(hListView);
        ListView_InsertItem(hListView, &lvItem);

        // Set window title in the second column
        ListView_SetItemText(hListView, lvItem.iItem, 1, const_cast<LPWSTR>(info.windowTitle.c_str()));

        // Set window handle in the third column
        ListView_SetItemText(hListView, lvItem.iItem, 2, handleStr);
    }

    // Optionally, you can auto-size columns after refreshing
    ListView_SetColumnWidth(hListView, 0, LVSCW_AUTOSIZE);
    ListView_SetColumnWidth(hListView, 1, LVSCW_AUTOSIZE_USEHEADER);
    ListView_SetColumnWidth(hListView, 2, LVSCW_AUTOSIZE_USEHEADER);
}

// Function to move selected item up or down
void MoveSelectedItem(int direction)
{
    int selected = ListView_GetNextItem(hListView, -1, LVNI_SELECTED);
    if (selected == -1)
        return;

    int newIndex = selected + direction;
    if (newIndex < 0 || newIndex >= (int)windowList.size())
        return;

    // Swap items in the vector
    std::swap(windowList[selected], windowList[newIndex]);

    // Update index numbers
    for (size_t i = 0; i < windowList.size(); ++i)
    {
        windowList[i].index = static_cast<int>(i) + 1;
    }

    // Refresh the list
    RefreshWindowList();

    // Reselect the moved item
    ListView_SetItemState(hListView, newIndex, LVIS_SELECTED, LVIS_SELECTED);
    ListView_EnsureVisible(hListView, newIndex, FALSE);
}

// ListView window procedure
LRESULT CALLBACK ListViewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    if (msg == WM_LBUTTONDBLCLK)
    {
        int selected = ListView_GetNextItem(hwnd, -1, LVNI_SELECTED);
        if (selected != -1)
        {
            LVITEM lvItem = { 0 };
            lvItem.mask = LVIF_PARAM;
            lvItem.iItem = selected;
            ListView_GetItem(hwnd, &lvItem);
            HWND hWnd = (HWND)lvItem.lParam;
            if (IsWindow(hWnd))
            {
                // Bring the window to the foreground and give it focus
                SetForegroundWindow(hWnd);
            }
        }
    }
    return CallWindowProc(OldListViewProc, hwnd, msg, wp, lp);
}

// Function to adjust control positions and sizes
void AdjustControls()
{
    RECT rcClient;
    GetClientRect(hMainWindow, &rcClient);
    int clientWidth = rcClient.right - rcClient.left;
    int clientHeight = rcClient.bottom - rcClient.top;

    // Minimum window size
    if (clientWidth < MIN_WINDOW_WIDTH)
        clientWidth = MIN_WINDOW_WIDTH;
    if (clientHeight < MIN_WINDOW_HEIGHT)
        clientHeight = MIN_WINDOW_HEIGHT;

    int x = MARGIN;
    int y = MARGIN;

    // Arrange buttons on the top
    SetWindowPos(hCaptureButton, NULL, x, y, BUTTON_WIDTH, BUTTON_HEIGHT, SWP_NOZORDER);
    x += BUTTON_WIDTH + MARGIN;

    SetWindowPos(hRestoreButton, NULL, x, y, SMALL_BUTTON_WIDTH, BUTTON_HEIGHT, SWP_NOZORDER);
    x += SMALL_BUTTON_WIDTH + MARGIN;

    SetWindowPos(hClearButton, NULL, x, y, SMALL_BUTTON_WIDTH, BUTTON_HEIGHT, SWP_NOZORDER);
    x += SMALL_BUTTON_WIDTH + MARGIN;

    SetWindowPos(hArrangeButton, NULL, x, y, BUTTON_WIDTH, BUTTON_HEIGHT, SWP_NOZORDER);
    x += BUTTON_WIDTH + MARGIN;

    // *** Position the new "Capture by Title" button ***
    SetWindowPos(hCaptureByTitleButton, NULL, x, y, BUTTON_WIDTH, BUTTON_HEIGHT, SWP_NOZORDER);
    x += BUTTON_WIDTH + MARGIN;

    // Reset x and move y down for next row
    x = MARGIN;
    y += BUTTON_HEIGHT + MARGIN;

    // Number of windows label and edit
    SetWindowPos(hNumWindowsLabel, NULL, x, y + 3, LABEL_WIDTH, LABEL_HEIGHT, SWP_NOZORDER);
    x += LABEL_WIDTH + MARGIN;

    SetWindowPos(hNumWindowsEdit, NULL, x, y, EDIT_WIDTH, LABEL_HEIGHT, SWP_NOZORDER);
    x += EDIT_WIDTH + MARGIN;

    // Monitor selection label and combo box
    SetWindowPos(hMonitorLabel, NULL, x, y + 3, LABEL_WIDTH, LABEL_HEIGHT, SWP_NOZORDER);
    x += LABEL_WIDTH + MARGIN;

    SetWindowPos(hMonitorComboBox, NULL, x, y, COMBOBOX_WIDTH, COMBOBOX_HEIGHT, SWP_NOZORDER);
    x += COMBOBOX_WIDTH + MARGIN;

    // Pixel Fix X label and edit box
    SetWindowPos(hPixelFixXLabel, NULL, x, y + 3, 150, LABEL_HEIGHT, SWP_NOZORDER);
    x += 150 + MARGIN;

    SetWindowPos(hPixelFixXEdit, NULL, x, y, 100, LABEL_HEIGHT, SWP_NOZORDER);
    x += 100 + MARGIN;

    // Pixel Fix Y label and edit box
    SetWindowPos(hPixelFixYLabel, NULL, x, y + 3, 150, LABEL_HEIGHT, SWP_NOZORDER);
    x += 150 + MARGIN;

    SetWindowPos(hPixelFixYEdit, NULL, x, y, 100, LABEL_HEIGHT, SWP_NOZORDER);
    x += 100 + MARGIN;

    // Min Vertical Spacing Y label and edit box
    SetWindowPos(hMinSpacingYLabel, NULL, x, y + 3, 200, LABEL_HEIGHT, SWP_NOZORDER);
    x += 200 + MARGIN;

    SetWindowPos(hMinSpacingYEdit, NULL, x, y, 100, LABEL_HEIGHT, SWP_NOZORDER);
    x += 100 + MARGIN;

    // *** Position the new Window Title label and edit box ***
    // Reset x and move y down for next row
    x = MARGIN;
    y += LABEL_HEIGHT + MARGIN;

    // Window Title label
    SetWindowPos(hWindowTitleLabel, NULL, x, y + 3, LABEL_WIDTH, LABEL_HEIGHT, SWP_NOZORDER);
    x += LABEL_WIDTH + MARGIN;

    // Window Title edit box
    SetWindowPos(hWindowTitleEdit, NULL, x, y, 200, LABEL_HEIGHT, SWP_NOZORDER);
    x += 200 + MARGIN;

    // *** Position the "Capture by Title" button ***
    SetWindowPos(hCaptureByTitleButton, NULL, x, y, BUTTON_WIDTH, BUTTON_HEIGHT, SWP_NOZORDER);
    x += BUTTON_WIDTH + MARGIN;

    // Adjust y for the next row
    y += LABEL_HEIGHT + MARGIN;

    // Adjust ListView size
    int listViewWidth = clientWidth - 3 * MARGIN - MOVE_BUTTON_WIDTH;
    int listViewHeight = clientHeight - y - MARGIN;
    SetWindowPos(hListView, NULL, MARGIN, y, listViewWidth, listViewHeight, SWP_NOZORDER);

    // Move Up and Move Down buttons
    int moveButtonX = MARGIN + listViewWidth + MARGIN;
    int moveButtonY = y;
    SetWindowPos(hMoveUpButton, NULL, moveButtonX, moveButtonY, MOVE_BUTTON_WIDTH, MOVE_BUTTON_HEIGHT, SWP_NOZORDER);
    moveButtonY += MOVE_BUTTON_HEIGHT + MARGIN;
    SetWindowPos(hMoveDownButton, NULL, moveButtonX, moveButtonY, MOVE_BUTTON_WIDTH, MOVE_BUTTON_HEIGHT, SWP_NOZORDER);

    // Adjust ListView column widths
    int totalColumnWidth = listViewWidth - GetSystemMetrics(SM_CXVSCROLL);
    int indexColWidth = 50;
    int handleColWidth = 120;
    int titleColWidth = totalColumnWidth - indexColWidth - handleColWidth;

    ListView_SetColumnWidth(hListView, 0, indexColWidth);
    ListView_SetColumnWidth(hListView, 1, titleColWidth);
    ListView_SetColumnWidth(hListView, 2, handleColWidth);
}

// Message box timer callback
VOID CALLBACK MsgBoxTimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
    if (g_hMsgBoxWnd)
    {
        PostMessage(g_hMsgBoxWnd, WM_CLOSE, 0, 0);
    }
    if (g_msgboxTimerId)
    {
        KillTimer(NULL, g_msgboxTimerId);
        g_msgboxTimerId = 0;
    }
}

// CBT hook procedure for message box timeout
LRESULT CALLBACK CBTProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HCBT_ACTIVATE)
    {
        // Get the handle of the message box window
        g_hMsgBoxWnd = (HWND)wParam;

        // Set the timer when the message box is activated
        if (g_msgboxTimerId)
        {
            KillTimer(NULL, g_msgboxTimerId);
        }
        g_msgboxTimerId = SetTimer(NULL, 0, 2000, MsgBoxTimerProc); // 2000 ms timeout

        // Unhook the CBT hook
        if (g_hMsgBoxHook)
        {
            UnhookWindowsHookEx(g_hMsgBoxHook);
            g_hMsgBoxHook = NULL;
        }
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

// Function to display a timed message box
int TimedMessageBox(HWND hWnd, LPCTSTR lpText, LPCTSTR lpCaption, UINT uType)
{
    // Reset the message box window handle
    g_hMsgBoxWnd = NULL;

    // Set the CBT hook
    g_hMsgBoxHook = SetWindowsHookEx(WH_CBT, CBTProc, NULL, GetCurrentThreadId());

    int ret = MessageBox(hWnd, lpText, lpCaption, uType);

    // Unhook the CBT hook in case it wasn't unhooked
    if (g_hMsgBoxHook)
    {
        UnhookWindowsHookEx(g_hMsgBoxHook);
        g_hMsgBoxHook = NULL;
    }
    // Kill the timer if it's still running
    if (g_msgboxTimerId)
    {
        KillTimer(NULL, g_msgboxTimerId);
        g_msgboxTimerId = 0;
    }

    return ret;
}

// Callback function for the main window
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
    {
        hMainWindow = hWnd;

        INITCOMMONCONTROLSEX icex = { sizeof(INITCOMMONCONTROLSEX), ICC_LISTVIEW_CLASSES };
        InitCommonControlsEx(&icex);

        hGlobalKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, GlobalKeyboardProc, hInstance, 0);
        if (!hGlobalKeyboardHook)
        {
            MessageBox(hWnd, L"Failed to set global keyboard hook.", L"Error", MB_OK | MB_ICONERROR);
        }
        // Create buttons
        hCaptureButton = CreateWindow(L"BUTTON", L"Capture Windows", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            0, 0, 0, 0, hWnd, (HMENU)ID_CAPTURE_BUTTON, NULL, NULL);

        hRestoreButton = CreateWindow(L"BUTTON", L"Restore", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            0, 0, 0, 0, hWnd, (HMENU)ID_RESTORE_BUTTON, NULL, NULL);

        hClearButton = CreateWindow(L"BUTTON", L"Clear", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            0, 0, 0, 0, hWnd, (HMENU)ID_CLEAR_BUTTON, NULL, NULL);

        hArrangeButton = CreateWindow(L"BUTTON", L"Arrange Windows", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            0, 0, 0, 0, hWnd, (HMENU)ID_ARRANGE_BUTTON, NULL, NULL);

        // *** Create the new "Capture by Title" button ***
        hCaptureByTitleButton = CreateWindow(L"BUTTON", L"Capture by Title", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            0, 0, 0, 0, hWnd, (HMENU)ID_CAPTURE_BY_TITLE_BUTTON, NULL, NULL);

        // Input field for number of windows
        hNumWindowsEdit = CreateWindow(L"EDIT", L"1", WS_TABSTOP | WS_VISIBLE | WS_CHILD | ES_NUMBER | WS_BORDER,
            0, 0, 0, 0, hWnd, (HMENU)ID_NUM_WINDOWS_EDIT, NULL, NULL);

        hNumWindowsLabel = CreateWindow(L"STATIC", L"Number of windows to capture:", WS_VISIBLE | WS_CHILD,
            0, 0, 0, 0, hWnd, NULL, NULL, NULL);

        // Monitor selection label and combo box
        hMonitorLabel = CreateWindow(L"STATIC", L"Select Monitor:", WS_VISIBLE | WS_CHILD,
            0, 0, 0, 0, hWnd, NULL, NULL, NULL);

        hMonitorComboBox = CreateWindow(L"COMBOBOX", NULL, CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL,
            0, 0, 0, 0, hWnd, (HMENU)ID_MONITOR_COMBOBOX, NULL, NULL);

        // Create Pixel Fix X Label and Edit Box
        hPixelFixXLabel = CreateWindow(L"STATIC", L"Pixel Fix X (±):", WS_VISIBLE | WS_CHILD,
            0, 0, 150, LABEL_HEIGHT, hWnd, (HMENU)ID_PIXEL_FIX_X_LABEL, NULL, NULL);

        hPixelFixXEdit = CreateWindow(L"EDIT", L"0", WS_TABSTOP | WS_VISIBLE | WS_CHILD | WS_BORDER,
            0, 0, 100, LABEL_HEIGHT, hWnd, (HMENU)ID_PIXEL_FIX_X_EDIT, NULL, NULL);

        // Create Pixel Fix Y Label and Edit Box
        hPixelFixYLabel = CreateWindow(L"STATIC", L"Pixel Fix Y (±):", WS_VISIBLE | WS_CHILD,
            0, 0, 150, LABEL_HEIGHT, hWnd, (HMENU)ID_PIXEL_FIX_Y_LABEL, NULL, NULL);

        hPixelFixYEdit = CreateWindow(L"EDIT", L"0", WS_TABSTOP | WS_VISIBLE | WS_CHILD | WS_BORDER,
            0, 0, 100, LABEL_HEIGHT, hWnd, (HMENU)ID_PIXEL_FIX_Y_EDIT, NULL, NULL);

        // Create Min Vertical Spacing Y Label and Edit Box
        hMinSpacingYLabel = CreateWindow(L"STATIC", L"Min Vertical Spacing (pixels):", WS_VISIBLE | WS_CHILD,
            0, 0, 200, LABEL_HEIGHT, hWnd, (HMENU)ID_MIN_SPACING_Y_LABEL, NULL, NULL);

        hMinSpacingYEdit = CreateWindow(L"EDIT", L"20", WS_TABSTOP | WS_VISIBLE | WS_CHILD | ES_NUMBER | WS_BORDER,
            0, 0, 100, LABEL_HEIGHT, hWnd, (HMENU)ID_MIN_SPACING_Y_EDIT, NULL, NULL);

        // *** Create the new Window Title label and edit box ***
        hWindowTitleLabel = CreateWindow(L"STATIC", L"Window Title:", WS_VISIBLE | WS_CHILD,
            0, 0, 100, LABEL_HEIGHT, hWnd, (HMENU)ID_WINDOWTITLE_LABEL, NULL, NULL);

        hWindowTitleEdit = CreateWindow(L"EDIT", L"", WS_TABSTOP | WS_VISIBLE | WS_CHILD | WS_BORDER,
            0, 0, 200, LABEL_HEIGHT, hWnd, (HMENU)ID_WINDOWTITLE_EDIT, NULL, NULL);

        // Update Monitor ComboBox
        UpdateMonitorComboBox();

        // ListView to display captured windows
        hListView = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL | WS_BORDER,
            0, 0, 0, 0, hWnd, (HMENU)ID_WINDOW_LISTVIEW, NULL, NULL);

        // Add columns to ListView
        LVCOLUMN lvCol = { 0 };
        lvCol.mask = LVCF_WIDTH | LVCF_TEXT;

        // Index column
        lvCol.cx = 50;
        lvCol.pszText = (LPWSTR)L"Index";
        ListView_InsertColumn(hListView, 0, &lvCol);

        // Window Title column
        lvCol.cx = 300; // Will be adjusted in AdjustControls
        lvCol.pszText = (LPWSTR)L"Window Title";
        ListView_InsertColumn(hListView, 1, &lvCol);

        // Handle column
        lvCol.cx = 120;
        lvCol.pszText = (LPWSTR)L"Handle";
        ListView_InsertColumn(hListView, 2, &lvCol);

        // Move Up and Move Down buttons
        hMoveUpButton = CreateWindow(L"BUTTON", L"Move Up", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            0, 0, 0, 0, hWnd, (HMENU)ID_MOVE_UP_BUTTON, NULL, NULL);

        hMoveDownButton = CreateWindow(L"BUTTON", L"Move Down", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            0, 0, 0, 0, hWnd, (HMENU)ID_MOVE_DOWN_BUTTON, NULL, NULL);

        // Set extended ListView styles
        ListView_SetExtendedListViewStyle(hListView, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

        // Subclass ListView to handle item activation
        OldListViewProc = (WNDPROC)SetWindowLongPtr(hListView, GWLP_WNDPROC, (LONG_PTR)ListViewProc);

        // Adjust initial control positions
        AdjustControls();

        // Set minimum window size
        SetWindowPos(hWnd, NULL, 0, 0, MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT, SWP_NOMOVE | SWP_NOZORDER);
    }
    break;
    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        switch (wmId)
        {
        case ID_CAPTURE_BUTTON: // Capture Windows
            StartCapturingWindows(hWnd);
            break;
        case ID_RESTORE_BUTTON: // Restore
            RestoreWindows();
            break;
        case ID_CLEAR_BUTTON: // Clear
            ClearCapturedWindows();
            break;
        case ID_ARRANGE_BUTTON: // Arrange Windows
            ArrangeWindows();
            break;
        case ID_MOVE_UP_BUTTON: // Move Up
            MoveSelectedItem(-1);
            break;
        case ID_MOVE_DOWN_BUTTON: // Move Down
            MoveSelectedItem(1);
            break;
        case ID_CAPTURE_BY_TITLE_BUTTON: // *** Handle "Capture by Title" Button ***
        {
            wchar_t titleBuffer[256];
            GetWindowText(hWindowTitleEdit, titleBuffer, 256);
            std::wstring windowTitle = titleBuffer;
            CaptureWindowsByTitle(windowTitle);
        }
        break;
        default:
            break;
        }
    }
    break;
    case WM_UNHOOK_HOOKS:
    {
        // Unhook the mouse and keyboard hooks
        if (hMouseHook)
        {
            UnhookWindowsHookEx(hMouseHook);
            hMouseHook = NULL;
        }
        if (hKeyboardHook)
        {
            UnhookWindowsHookEx(hKeyboardHook);
            hKeyboardHook = NULL;
        }

        // Reset flags
        isCapturing = false;
        bool wasCaptureCompleted = captureCompleted;
        captureCompleted = false;

        // Display the appropriate MessageBox
        if (windowsCaptured >= windowsToCapture && wasCaptureCompleted)
        {
            TimedMessageBox(NULL, L"All windows have been captured.", L"Info", MB_OK);
        }
        else
        {
            TimedMessageBox(NULL, L"Window capturing has been canceled.", L"Info", MB_OK);
        }
    }
    break;
    case WM_SIZE:
        AdjustControls();
        break;
    case WM_DESTROY:
        if (hGlobalKeyboardHook)
        {
            UnhookWindowsHookEx(hGlobalKeyboardHook);
            hGlobalKeyboardHook = NULL;
        }
        if (hMouseHook)
        {
            UnhookWindowsHookEx(hMouseHook);
            hMouseHook = NULL;
        }
        if (hKeyboardHook)
        {
            UnhookWindowsHookEx(hKeyboardHook);
            hKeyboardHook = NULL;
        }
        // Kill the timer if it's running
        if (g_msgboxTimerId)
        {
            KillTimer(NULL, g_msgboxTimerId);
            g_msgboxTimerId = 0;
        }
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// *** New: wWinMain remains the same ***
int APIENTRY wWinMain(_In_ HINSTANCE hInst,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR    lpCmdLine,
    _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    hInstance = hInst;

    // Register window class
    WNDCLASSEX wcex = { 0 };
    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = 0;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszMenuName = NULL;
    wcex.lpszClassName = L"MyWindowClass";
    wcex.hIconSm = LoadIcon(NULL, IDI_APPLICATION);

    if (!RegisterClassEx(&wcex))
    {
        MessageBox(NULL, L"Window class registration failed.", L"Error", MB_OK);
        return 1;
    }

    // Create main window
    HWND hWnd = CreateWindow(L"MyWindowClass", L"Window Management Tool",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, 0, MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT, NULL, NULL, hInstance, NULL);

    if (!hWnd)
    {
        MessageBox(NULL, L"Main window creation failed.", L"Error", MB_OK);
        return 1;
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    // Message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}
