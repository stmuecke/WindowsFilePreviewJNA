/*
 * Copyright (c) 2025 by Stefan Mücke
 * 
 * Permission to use, copy, modify, and/or distribute this software
 * for any purpose with or without fee is hereby granted.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
 * WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
 * AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT, INDIRECT,
 * OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE,
 * DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION,
 * ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
 * 
 * SPDX-License-Identifier: MIT-0
 */
package stm.ui.win32;

import java.awt.Canvas;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.Insets;
import java.awt.Rectangle;
import java.awt.RenderingHints;
import java.awt.TexturePaint;
import java.awt.Window;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.image.BufferStrategy;
import java.awt.image.BufferedImage;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.util.EventListener;
import java.util.Queue;
import java.util.ResourceBundle;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.LinkedBlockingQueue;

import javax.swing.JLabel;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.event.EventListenerList;

import com.sun.jna.Memory;
import com.sun.jna.Native;
import com.sun.jna.Pointer;
import com.sun.jna.Structure;
import com.sun.jna.Structure.FieldOrder;
import com.sun.jna.WString;
import com.sun.jna.platform.win32.GDI32;
import com.sun.jna.platform.win32.Guid.CLSID;
import com.sun.jna.platform.win32.Guid.IID;
import com.sun.jna.platform.win32.Guid.REFIID;
import com.sun.jna.platform.win32.Ole32;
import com.sun.jna.platform.win32.ShellAPI.SHELLEXECUTEINFO;
import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WTypes;
import com.sun.jna.platform.win32.WinDef.HBITMAP;
import com.sun.jna.platform.win32.WinDef.HDC;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinDef.RECT;
import com.sun.jna.platform.win32.WinError;
import com.sun.jna.platform.win32.WinGDI;
import com.sun.jna.platform.win32.WinGDI.BITMAP;
import com.sun.jna.platform.win32.WinGDI.BITMAPINFO;
import com.sun.jna.platform.win32.WinNT.HANDLE;
import com.sun.jna.platform.win32.WinNT.HRESULT;
import com.sun.jna.platform.win32.WinUser.MSG;
import com.sun.jna.platform.win32.WinUser.SIZE;
import com.sun.jna.platform.win32.COM.COMUtils;
import com.sun.jna.platform.win32.COM.Unknown;
import com.sun.jna.ptr.IntByReference;
import com.sun.jna.ptr.PointerByReference;
import com.sun.jna.win32.StdCallLibrary;
import com.sun.jna.win32.W32APIOptions;

/**
 * Shows native file previews on Windows using {@code IPreviewHandler} or {@code IThumbnailProvider} as a fallback.
 * <p>
 * Supports {@code IInitializeWithStream}, {@code IInitializeWithFile}, and {@code IInitializeWithItem} for handler input.
 * 
 * @see <a href="https://learn.microsoft.com/en-us/windows/win32/shell/preview-handlers">Preview Handlers and Shell Preview Host</a>
 * @see <a href="https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ipreviewhandler">IPreviewHandler</a>
 * @see <a href="https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishellitemimagefactory">IShellItemImageFactory</a>
 * @see <a href="https://learn.microsoft.com/en-us/windows/win32/api/thumbcache/nn-thumbcache-ithumbnailprovider">IThumbnailProvider</a>
 * @see <a href="https://geelaw.blog/entries/ipreviewhandlerframe-wpf-1-ui-assoc/">Hosting a preview handler in WPF, part 1: UI and file associations</a>
 * @see <a href="https://geelaw.blog/entries/ipreviewhandlerframe-wpf-2-interop/">Hosting a preview handler in WPF, correctly, part 2: interop</a>
 */
public class WindowsPreviewCanvas extends Canvas {

	public interface PreviewListener extends EventListener {
		void onPreviewLoaded(PreviewInfo info);
	}

	/**
	 * Contains messages from the ResourceBundle. Values may be customized at runtime.
	 */
	public static class PreviewMessages {
		public String emptyFile;
		public String fileNotFound;
		public String initialMessage;
		public String nonWindowsPlatform;
		public String noPreviewAvailable;
	}

	/**
	 * Contains information about the current preview or thumbnail.
	 */
	public static class PreviewInfo implements Cloneable {

		public File file;
		public String fileExtension;
		public String interfaceType;
		public String iid;
		public String clsid;
		public String handlerName;
		public String initializerType;

		@Override
		public PreviewInfo clone() {
			try {
				return (PreviewInfo) super.clone();
			} catch (CloneNotSupportedException e) {
				throw new RuntimeException();
			}
		}
	}

	interface Shlwapi extends StdCallLibrary {

		Shlwapi INSTANCE = Native.load("Shlwapi", Shlwapi.class, W32APIOptions.DEFAULT_OPTIONS);

		int STGM_READ = 0;
		int ASSOCF_INIT_DEFAULTTOSTAR = 4;
		int ASSOCSTR_SHELLEXTENSION = 16;

		HRESULT AssocQueryStringW(int flags, int str, WString pszAssoc, WString pszExtra, char[] pszOut, int[] pcchOut);

		HRESULT SHCreateStreamOnFileEx(WString pszFile, int grfMode, int dwAttributes, boolean fCreate, Pointer pstmTemplate, PointerByReference ppstm);

	}

	interface Shell32 extends com.sun.jna.platform.win32.Shell32 {

		Shell32 INSTANCE = Native.load("shell32", Shell32.class, W32APIOptions.DEFAULT_OPTIONS);

		int SW_SHOWNORMAL = 1;
		int SW_SHOWDEFAULT = 10;

		HRESULT SHCreateItemFromParsingName(WString pszPath, Pointer pbc, REFIID riid, PointerByReference ppv);

	}

	/**
	 * @see <a href="https://learn.microsoft.com/en-us/windows/win32/api/shtypes/ns-shtypes-logfontw">LOGFONTW structure (shtypes.h)</a>
	 */
	@FieldOrder({"lfHeight", "lfWidth", "lfEscapement", "lfOrientation", "lfWeight", "lfItalic", "lfUnderline", "lfStrikeOut", "lfCharSet", "lfOutPrecision", "lfClipPrecision", "lfQuality", "lfPitchAndFamily", "lfFaceName"})
	public static class LOGFONT extends Structure {

		public int lfHeight;
		public int lfWidth;
		public int lfEscapement;
		public int lfOrientation;
		public int lfWeight;
		public byte lfItalic;
		public byte lfUnderline;
		public byte lfStrikeOut;
		public byte lfCharSet;
		public byte lfOutPrecision;
		public byte lfClipPrecision;
		public byte lfQuality;
		public byte lfPitchAndFamily;
		public char[] lfFaceName = new char[32];

	}

	static class SIZEByValue extends SIZE implements Structure.ByValue {

		public SIZEByValue(int width, int height) {
			super(width, height);
		}

	}

	/**
	 * Wrapper for {@code IPreviewHandler}.
	 */
	static class PreviewHandler extends Unknown {

		public PreviewHandler(Pointer pvInstance) {
			super(pvInstance);
		}

		public HRESULT SetWindow(HWND hwnd, RECT prc) {
			return new HRESULT(_invokeNativeInt(3, new Object[] {getPointer(), hwnd, prc}));
		}

		public HRESULT SetRect(RECT prc) {
			return new HRESULT(_invokeNativeInt(4, new Object[] {getPointer(), prc}));
		}

		public HRESULT DoPreview() {
			return new HRESULT(_invokeNativeInt(5, new Object[] {getPointer()}));
		}

		public HRESULT Unload() {
			return new HRESULT(_invokeNativeInt(6, new Object[] {getPointer()}));
		}

		public HRESULT SetFocus() {
			return new HRESULT(_invokeNativeInt(7, new Object[] {getPointer()}));
		}

		public HRESULT QueryFocus(PointerByReference phwnd) {
			return new HRESULT(_invokeNativeInt(8, new Object[] {getPointer(), phwnd}));
		}

		public HRESULT TranslateAccelerator(MSG msg) {
			return new HRESULT(_invokeNativeInt(9, new Object[] {getPointer(), msg}));
		}

	}

	/**
	 * Wrapper for {@code IPreviewHandlerVisuals}.
	 */
	public static class PreviewHandlerVisuals extends Unknown {

		public PreviewHandlerVisuals(Pointer pvInstance) {
			super(pvInstance);
		}

		public HRESULT SetBackgroundColor(int color) {
			return new HRESULT(_invokeNativeInt(3, new Object[] {getPointer(), color}));
		}

		public HRESULT SetFont(LOGFONT plf) {
			return new HRESULT(_invokeNativeInt(4, new Object[] {getPointer(), plf.getPointer()}));
		}

		public HRESULT SetTextColor(int color) {
			return new HRESULT(_invokeNativeInt(5, new Object[] {getPointer(), color}));
		}

	}

	/**
	 * Wrapper for {@code IInitializeWithFile}.
	 */
	static class InitializeWithFile extends Unknown {

		public InitializeWithFile(Pointer pvInstance) {
			super(pvInstance);
		}

		public HRESULT Initialize(WString pszFilePath, int grfMode) {
			return new HRESULT(_invokeNativeInt(3, new Object[] {getPointer(), pszFilePath, grfMode}));
		}

	}

	/**
	 * Wrapper for {@code IInitializeWithStream}.
	 */
	static class InitializeWithStream extends Unknown {

		public InitializeWithStream(Pointer pvInstance) {
			super(pvInstance);
		}

		public HRESULT Initialize(Pointer pStream, int grfMode) {
			return new HRESULT(_invokeNativeInt(3, new Object[] {getPointer(), pStream, grfMode}));
		}

	}

	/**
	 * Wrapper for {@code IInitializeWithItem}.
	 */
	static class InitializeWithItem extends Unknown {

		public InitializeWithItem(Pointer pvInstance) {
			super(pvInstance);
		}

		public HRESULT Initialize(Pointer pShellItem, int grfMode) {
			return new HRESULT(_invokeNativeInt(3, new Object[] {getPointer(), pShellItem, grfMode}));
		}

	}

	/**
	 * Wrapper for {@code IThumbnailProvider}.
	 */
	static class ThumbnailProvider extends Unknown {

		public ThumbnailProvider(Pointer pvInstance) {
			super(pvInstance);
		}

		public HRESULT GetThumbnail(int cx, PointerByReference phbmp, IntByReference pdwAlpha) {
			return new HRESULT(_invokeNativeInt(3, new Object[] {getPointer(), cx, phbmp.getPointer(), pdwAlpha.getPointer()}));
		}

	}

	/**
	 * Wrapper for {@code IShellItemImageFactory}.
	 */
	static class ShellItemImageFactory extends Unknown {

		public ShellItemImageFactory(Pointer pvInstance) {
			super(pvInstance);
		}

		public HRESULT GetImage(SIZE size, int flags, PointerByReference phbm) {
			return new HRESULT(_invokeNativeInt(3, new Object[] {getPointer(), size, flags, phbm.getPointer()}));
		}

	}

	static class Preview {
		PreviewThread thread;
		Unknown handler;
		boolean isPreviewMode;
		boolean isThumbnailMode;
		/** Flag that a larger thumbnail is not available */
		boolean isBiggestThumbnail;
		boolean isInitialized;
		boolean hasVisualsInterface;
		volatile boolean isCanceled;
		BufferedImage thumbnailImage;
		Throwable exception;
		long unloadStartMillis;
		PreviewInfo info = new PreviewInfo();
	}

	static class PreviewThread extends Thread {

		private final BlockingQueue<Runnable> queue = new LinkedBlockingQueue<>();
		private volatile boolean keepRunning = true;

		@Override
		public void run() {
			while (keepRunning) {
				try {
					Runnable task = queue.take();
					task.run();
				} catch (InterruptedException e) {
					System.out.println("PreviewThread: interrupted");
					// TODO Anything to do here?
				} catch (Throwable t) {
					t.printStackTrace();
				}
			}
		}

		public void submit(Runnable task) {
			try {
				queue.put(task);
			} catch (InterruptedException e) {
				Thread.currentThread().interrupt();
			}
		}

		public void shutdown() {
			keepRunning = false;
			interrupt();
		}

	}
	
	class Watchdog extends Thread {
		
		boolean keepRunning = true;
		
		@SuppressWarnings("deprecation")
		@Override
		public void run() {
			while (keepRunning || unloadQueue.size() > 0) {
				for (Preview preview : unloadQueue) {
					if (!preview.thread.isAlive()) {
						unloadQueue.remove(preview);
					} else {
						long elapsed = System.currentTimeMillis() - preview.unloadStartMillis;
						if (elapsed > 60000) {
							preview.thread.stop(); // TODO No longer available on Java 23+
							unloadQueue.remove(preview);
						}
					}
				}
				try {
					Thread.sleep(10000); 
				} catch (InterruptedException e) {
					Thread.currentThread().interrupt();
				}
			}
		}
		
	}

	private static final IID IID_IPreviewHandler = new IID("{8895b1c6-b41f-4c1c-a562-0d564250836f}");
	private static final IID IID_IPreviewHandlerVisuals = new IID("{196bf9a5-b346-4ef0-aa1e-5dcdb76768b1}");
	private static final IID IID_IThumbnailProvider = new IID("{e357fccd-a995-4576-b01f-234630154e96}");
	private static final IID IID_IShellItemImageFactory = new IID("{bcc18b79-ba16-442f-80c4-8a59c30c463b}");
	private static final IID IID_IInitializeWithFile = new IID("{b7d14566-0509-4cce-a71f-0a554233bd9b}");
	private static final IID IID_IInitializeWithStream = new IID("{b824b49d-22ac-4161-ac8a-9916e8fa3f7f}");
	private static final IID IID_IInitializeWithItem = new IID("{7f73be3f-fb79-493c-a6c7-7ee14e245841}");
	private static final IID IID_IShellItem = new IID("{43826d1e-e718-42ee-bc55-a1e261c37bfe}");

	private static final int WTSAT_UNKNOWN = 0;
	private static final int WTSAT_RGB = 1;
	private static final int WTSAT_ARGB = 2;

	private static final int SIIGBF_BIGGERSIZEOK = 1;
	private static final int SIIGBF_THUMBNAILONLY = 8;

	// UI
	private HWND canvasHwnd;
	private JLabel messageLabel = new JLabel();
	private WindowAdapter windowListener;
	private PropertyChangeListener lafListener;
	private TexturePaint checkeredPaint;

	// Configuration
	private boolean isWindows = System.getProperty("os.name").startsWith("Windows");
	private Color previewBackgroundColor;
	private Color previewTextColor;
	private Color messageColor;
	private Font previewFont;
	private boolean openThumbnailOnDoubleClick = true;
	private boolean showCheckeredBackground = true;
	private Insets thumbnailInsets = new Insets(0, 0, 0, 0);
	private PreviewMessages messages;
	int checkeredTileSize = 8;

	// State
	private Window window;
	private String previewMessage;
	private File currFile;
	private Preview currPreview;
	private boolean isMouseOver;
	private Queue<Preview> unloadQueue = new ConcurrentLinkedQueue<>(); // for proper cleanup
	
	// Auxiliary
	private EventListenerList listenerList = new EventListenerList();
	private Watchdog watchdog;

	public WindowsPreviewCanvas() {
		System.setProperty("sun.awt.noerasebackground", "true"); // Avoid flickering
		previewBackgroundColor = UIManager.getColor("List.background");
		previewTextColor = UIManager.getColor("TextField.foreground");
		messageColor = UIManager.getColor("Label.foreground");
		addComponentListener(new ComponentAdapter() {
			@Override
			public void componentResized(ComponentEvent e) {
				updatePreviewRect();
			}
		});
		addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent me) {
				if (currPreview != null && currPreview.isThumbnailMode && openThumbnailOnDoubleClick && me.getClickCount() == 2) {
					openFile(WindowsPreviewCanvas.this, currFile);
				}
			}
			@Override
			public void mouseEntered(MouseEvent e) {
				isMouseOver = true;
				if (showCheckeredBackground && currPreview != null && currPreview.isThumbnailMode)
					repaintImmediately();
			}
			@Override
			public void mouseExited(MouseEvent e) {
				isMouseOver = false;
				if (showCheckeredBackground && currPreview != null && currPreview.isThumbnailMode)
					repaintImmediately();
			}
		});
		setMinimumSize(new Dimension(0, 0));

		// Load messages
		ResourceBundle bundle = null;
		try {
			bundle = ResourceBundle.getBundle(getClass().getName());
		} catch (Exception e) {
		}
		messages = new PreviewMessages();
		messages.emptyFile = getMessage(bundle, "initialMessage");
		messages.initialMessage = getMessage(bundle, "initialMessage");
		messages.nonWindowsPlatform = getMessage(bundle, "nonWindowsMessage");
		messages.fileNotFound = getMessage(bundle, "fileNotFound");
		messages.emptyFile = getMessage(bundle, "emptyFile");
		messages.noPreviewAvailable = getMessage(bundle, "noPreviewAvailable");

		// Set initial message
		previewMessage = isWindows ? messages.initialMessage : messages.nonWindowsPlatform;
	}

	private String getMessage(ResourceBundle bundle, String key) {
		try {
			String value = bundle == null ? null : bundle.getString(key);
			return value == null ? "%" + key + "%" : value;
		} catch (Exception e) {
			e.printStackTrace();
			return "%" + key + "%";
		}
	}

	public PreviewMessages getMessages() {
		return messages;
	}

	public void showPreview(File file) {
		this.currFile = file;
		if (!isDisplayable())
			return;
		if (!isWindows) {
			showPreviewMessage(messages.nonWindowsPlatform);
		} else if (file == null) {
			showPreviewMessage(messages.initialMessage);
		} else if (!file.exists()) {
			showPreviewMessage(messages.fileNotFound);
		} else if (file.length() == 0) {
			showPreviewMessage(messages.emptyFile);
		} else {
			hidePreview();

			// Create preview
			Preview preview = createPreview(file);
			if (preview == null) {
				showPreviewMessage(messages.noPreviewAvailable);
				return;
			}
			currPreview = preview;
			preview.thread.submit(() -> loadPreview(preview));
		}
	}

	private void showPreviewMessage(String message) {
		System.out.println("showPreviewMessage: " + message);
		hidePreview();
		previewMessage = message;
		repaint();
	}

	private void hidePreview() {
		unloadPreviewAsync(currPreview);
		currPreview = null;
	}

	private Preview createPreview(File file) {
		String path = file.getAbsolutePath();
		int dot = path.lastIndexOf('.');
		String ext = dot == -1 ? "*" : path.substring(dot);

		showPreviewMessage("Vorschau laden..."); // TODO Localize

		Preview preview = new Preview();
		preview.info.file = file;
		preview.info.fileExtension = ext;

		// Run COM code in a separate thread
		preview.thread = new PreviewThread();
		preview.thread.start();
		preview.thread.submit(() -> {
			Ole32.INSTANCE.CoInitializeEx(null, Ole32.COINIT_MULTITHREADED);
			try {
				char[] out = new char[40];
				int[] len = new int[] {out.length};

				/*
				 * Note: CoCreateInstance may occasionally fail. The workaround is to try again. For details see:
				 * https://blogs.msdn.microsoft.com/adioltean/2005/06/24/when-cocreateinstance-returns-0x80080005-co_e_server_exec_failure/ 
				 */
				
				// Try to create an IPreviewHandler 
				HRESULT hr = Shlwapi.INSTANCE.AssocQueryStringW(Shlwapi.ASSOCF_INIT_DEFAULTTOSTAR, Shlwapi.ASSOCSTR_SHELLEXTENSION, new WString(ext), new WString(IID_IPreviewHandler.toGuidString()), out, len);
				if (COMUtils.SUCCEEDED(hr)) {
					String clsid = new String(out, 0, len[0] - 1);
					PointerByReference p = new PointerByReference();
					hr = Ole32.INSTANCE.CoCreateInstance(new CLSID(clsid), null, WTypes.CLSCTX_LOCAL_SERVER, IID_IPreviewHandler, p);
					if (hr.intValue() == WinError.CO_E_SERVER_EXEC_FAILURE) {
						hr = Ole32.INSTANCE.CoCreateInstance(new CLSID(clsid), null, WTypes.CLSCTX_LOCAL_SERVER, IID_IPreviewHandler, p);
					}
					if (COMUtils.SUCCEEDED(hr)) {
						preview.isPreviewMode = true;
						preview.info.interfaceType = "IPreviewHandler";
						preview.info.iid = IID_IPreviewHandler.toGuidString();
						preview.info.clsid = clsid;
						preview.handler = new PreviewHandler(p.getValue());
					} else {
						// TODO Show error message
						// COMUtils.checkRC(hr);
						String msg = "Error with CLSID=" + clsid + ":  0x" + Integer.toHexString(hr.intValue());
						System.out.println(msg);
						showPreviewMessage(msg);
					}
					return;
				}

				// Try to create an IThumbnailProvider
				hr = Shlwapi.INSTANCE.AssocQueryStringW(Shlwapi.ASSOCF_INIT_DEFAULTTOSTAR, Shlwapi.ASSOCSTR_SHELLEXTENSION, new WString(ext), new WString(IID_IThumbnailProvider.toGuidString()), out, len);
				if (COMUtils.SUCCEEDED(hr)) {
					String clsid = new String(out, 0, len[0] - 1);
					PointerByReference p = new PointerByReference();
					hr = Ole32.INSTANCE.CoCreateInstance(new CLSID(clsid), null, WTypes.CLSCTX_LOCAL_SERVER, IID_IThumbnailProvider, p);
					if (hr.intValue() == WinError.CO_E_SERVER_EXEC_FAILURE) {
						hr = Ole32.INSTANCE.CoCreateInstance(new CLSID(clsid), null, WTypes.CLSCTX_LOCAL_SERVER, IID_IThumbnailProvider, p);
					}
					if (COMUtils.SUCCEEDED(hr)) {
						preview.isThumbnailMode = true;
						preview.info.interfaceType = "IThumbnailProvider";
						preview.info.iid = IID_IThumbnailProvider.toGuidString();
						preview.info.clsid = clsid;
						preview.handler = new ThumbnailProvider(p.getValue());
						return;
					}
				}

				// Try to get IShellItemImageFactory as a secondary fallback
				PointerByReference itemPtr = new PointerByReference();
				hr = Shell32.INSTANCE.SHCreateItemFromParsingName(new WString(path), null, new REFIID(IID_IShellItem), itemPtr);
				if (COMUtils.SUCCEEDED(hr)) {
					Unknown shellItem = new Unknown(itemPtr.getValue());
					PointerByReference factoryPtr = new PointerByReference();
					hr = shellItem.QueryInterface(new REFIID(IID_IShellItemImageFactory), factoryPtr);
					if (COMUtils.SUCCEEDED(hr)) {
						preview.isThumbnailMode = true;
						preview.isInitialized = true;
						preview.info.interfaceType = "IShellItemImageFactory";
						preview.info.iid = IID_IShellItemImageFactory.toGuidString();
						preview.info.clsid = "";
						preview.handler = new ShellItemImageFactory(factoryPtr.getValue());
					}
					shellItem.Release();
				}

				// Stop thread if no handler has been found
				if (preview.handler == null) {
					preview.thread.shutdown();
					if (!preview.isCanceled) {
						showPreviewMessage(messages.noPreviewAvailable);
					}
				}

			} catch (Throwable e) {
				preview.exception = e;
				showPreviewMessage("Error: " + e);
			}
		});

		return preview;
	}

	private void loadPreview(Preview preview) {
		canvasHwnd = new HWND(Native.getComponentPointer(WindowsPreviewCanvas.this));
		if (canvasHwnd == null)
			return;
		if (preview.isCanceled)
			return;
		try {
			if (preview.isPreviewMode) {
				PreviewHandler handler = (PreviewHandler) preview.handler;

				// Unload previous content
				if (preview.isInitialized) {
					handler.Unload();
				}
				if (preview.isCanceled)
					return;

				// Initialize with stream, file, or item
				initializeHandler(preview);
				preview.isInitialized = true;
				if (preview.isCanceled)
					return;

				// Set window
				RECT zeroRect = new RECT();
				zeroRect.write();
				handler.SetWindow(canvasHwnd, zeroRect);

				// Customize visuals
				PointerByReference visualsPtr = new PointerByReference();
				HRESULT hr = handler.QueryInterface(new REFIID(IID_IPreviewHandlerVisuals), visualsPtr);
				if (COMUtils.SUCCEEDED(hr)) {
					preview.hasVisualsInterface = true;
					PreviewHandlerVisuals visuals = new PreviewHandlerVisuals(visualsPtr.getValue());
					if (previewBackgroundColor != null)
						visuals.SetBackgroundColor(previewBackgroundColor.getRGB() & 0xffffff); // without alpha
					if (previewTextColor != null)
						visuals.SetTextColor(previewTextColor.getRGB() & 0xffffff); // without alpha
					if (previewFont != null) {
						LOGFONT lf = new LOGFONT();
						int pointSize = previewFont.getSize();
						lf.lfHeight = -pointSize;
						lf.lfWeight = previewFont.isBold() ? 700 : 400;
						lf.lfItalic = (byte) (previewFont.isItalic() ? 1 : 0);
						String fontName = previewFont.getFontName();
						System.arraycopy(fontName.toCharArray(), 0, lf.lfFaceName, 0, Math.min(fontName.length(), lf.lfFaceName.length - 1));
						lf.lfFaceName[Math.min(fontName.length(), lf.lfFaceName.length - 1)] = '\0';
						lf.write();
						visuals.SetFont(lf);
					}
					visuals.Release();
				}

				// Render preview
				if (preview.isCanceled)
					return;
				handler.DoPreview();

				// Set rect
				RECT rect = new RECT();
				User32.INSTANCE.GetClientRect(canvasHwnd, rect);
				handler.SetRect(rect);

			} else {
				// Determine thumbnail/image size
				Rectangle b = getBounds();
				int requestedSize = Math.max(256, Math.max(b.width, b.height));

				// Obtain thumbnail
				PointerByReference phbmpRef = new PointerByReference();
				HBITMAP hbmp = null;
				int alphaType = 0;
				if (preview.handler instanceof ThumbnailProvider) {
					ThumbnailProvider provider = (ThumbnailProvider) preview.handler;

					// Initialize with stream, file, or item
					if (!preview.isInitialized) {
						initializeHandler(preview);
						preview.isInitialized = true;
					}
					IntByReference pdwAlpha = new IntByReference();
					HRESULT hr = provider.GetThumbnail(requestedSize, phbmpRef, pdwAlpha);
					COMUtils.checkRC(hr);
					alphaType = pdwAlpha.getValue();
					hbmp = new HBITMAP(phbmpRef.getValue());

				} else if (preview.handler instanceof ShellItemImageFactory) {
					ShellItemImageFactory factory = (ShellItemImageFactory) preview.handler;
					/*
					 * Problem 1: IShellItemImageFactory seems to impose a limit on the maximum size.
					 * When the limit of 1280 x 1280 is exceeded, no image will be returned (Window 10).
					 * The workaround is to limit the requested size.
					 * 
					 * Problem 2: IShellItemImageFactory sometimes returns transparent pixels as black.
					 * Looks like a Windows bug, but the exact reason is currently unknown. For sizes > 768,
					 * the problem doesn't seem to occur. The workaround is to request a size of at least 769.
					 */
					requestedSize = clamp(requestedSize, 769, 1280);

					SIZEByValue size = new SIZEByValue(requestedSize, requestedSize);
					size.write();
					int flags = SIIGBF_BIGGERSIZEOK | SIIGBF_THUMBNAILONLY;
					HRESULT hr = factory.GetImage(size, flags, phbmpRef);
					if (COMUtils.SUCCEEDED(hr)) {
						alphaType = WTSAT_ARGB;
						hbmp = new HBITMAP(phbmpRef.getValue());
					} else {
						showPreviewMessage(messages.noPreviewAvailable);
					}
				} else {
					showPreviewMessage(messages.noPreviewAvailable);
					return;
				}

				// Convert HBITMAP to BufferedImage
				if (hbmp != null) {
					preview.thumbnailImage = hbitmapToBufferedImage(hbmp, alphaType);
					int returnedSize = Math.max(preview.thumbnailImage.getWidth(), preview.thumbnailImage.getHeight());
					preview.isBiggestThumbnail = returnedSize < requestedSize && false;
					GDI32.INSTANCE.DeleteObject(hbmp);
				}

				// Repaint
				if (preview.isCanceled)
					return;
				SwingUtilities.invokeLater(() -> repaint());

			}
			
			PreviewInfo info = getPreviewInfo();
			for (PreviewListener listener : listenerList.getListeners(PreviewListener.class)) {
				listener.onPreviewLoaded(info);
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			showPreviewMessage("Error: " + e); // TODO Localization
		}
	}

	private void unloadPreviewAsync(Preview preview) {
		if (preview == null)
			return;
		unloadQueue.add(preview);
		preview.isCanceled = true;
		preview.thread.submit(() -> {
			if (preview.handler != null) {
				try {
					if (preview.isPreviewMode) {
						PreviewHandler handler = (PreviewHandler) preview.handler;
						RECT zero = new RECT();
						zero.write();
						handler.SetRect(zero);
						SwingUtilities.invokeLater(() -> repaint());
						handler.Unload();
						handler.Release();
					} else if (preview.isThumbnailMode) {
						preview.isThumbnailMode = false;
						preview.thumbnailImage = null;
						preview.handler.Release();
					}
					Ole32.INSTANCE.CoUninitialize();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			preview.thread.shutdown();
			unloadQueue.remove(preview);
		});
	}

	/**
	 * Provides details about the current preview.
	 */
	public PreviewInfo getPreviewInfo() {
		return currPreview == null ? new PreviewInfo() : currPreview.info.clone();
	}

	public void setFocus() {
		Preview preview = currPreview;
		if (preview != null && preview.isPreviewMode) {
			preview.thread.submit(() -> ((PreviewHandler) preview.handler).SetFocus());
		} else {
			requestFocus();
		}
	}

	public void setInitialMessage(String initialMessage) {
		getMessages().initialMessage = initialMessage;
	}

	/**
	 * Sets the color of the message that is shown when no preview or thumbnail is shown.
	 */
	public void setMessageColor(Color color) {
		this.messageColor = color;
	}

	/**
	 * Sets whether double-click on thumbnails should open them with the default registered application. Default to true.
	 */
	public void setOpenThumbnailOnDoubleClick(boolean openOnDoubleClick) {
		this.openThumbnailOnDoubleClick = openOnDoubleClick;
	}

	/**
	 * Sets the background color to be used with {@code IPreviewHandlerVisuals.SetBackgroundColor}. 
	 */
	public void setPreviewBackgroundColor(Color color) {
		this.previewBackgroundColor = color == null ? UIManager.getColor("List.background") : color;
	}

	/**
	 * Sets the background color to be used with {@code IPreviewHandlerVisuals.SetFont}. 
	 */
	public void setPreviewFont(Font font) {
		this.previewFont = font;
	}

	/**
	 * Sets the background color to be used with {@code IPreviewHandlerVisuals.SetTextColor}. 
	 */
	public void setPreviewTextColor(Color color) {
		this.previewTextColor = color;
	}

	/**
	 * Sets the insets for use in thumbnail mode (with {@code IThumbnailProvider} and {@code IShellItemImageFactory}).
	 * <p>
	 * Note: The insets do not apply to previews with an {@code IPreviewHandler} because many preview handlers do not
	 * respect the bounds given to them via {@code SetWindowPos} and {@code SetRect} and always have to be at (0,0)
	 * relative to their Canvas parent window.
	 */
	public void setThumbnailInsets(Insets insets) {
		if (insets == null)
			insets = new Insets(0, 0, 0, 0);
		this.thumbnailInsets = insets;
		updatePreviewRect();
	}
	
	public void addPreviewListener(PreviewListener listener) {
		listenerList.add(PreviewListener.class, listener);
	}

	public void removePreviewListener(PreviewListener listener) {
		listenerList.remove(PreviewListener.class, listener);
	}
	
	@Override
	public void addNotify() {
		super.addNotify();
		start();
	}

	@Override
	public void removeNotify() {
		stop();
		super.removeNotify();
	}

	private void start() {
		installWindowListener();
		installLafListener();
		watchdog = new Watchdog();
		watchdog.start();
		if (currFile != null)
			showPreview(currFile);
	}

	private void stop() {
		if (window != null)
			window.removeWindowListener(windowListener);
		UIManager.removePropertyChangeListener(lafListener);
		hidePreview();
		watchdog.keepRunning = false;
		for (Preview preview : unloadQueue) {
			try {
				preview.thread.join(2000);
			} catch (InterruptedException e) {
				Thread.currentThread().interrupt();
			}
		}
	}

	private void installWindowListener() {
		window = SwingUtilities.getWindowAncestor(this);
		if (window == null)
			return;
		window.addWindowListener(windowListener = new WindowAdapter() {
			@Override
			public void windowClosing(WindowEvent e) {
				stop();
			}
		});
	}

	private void installLafListener() {
		UIManager.addPropertyChangeListener(lafListener = evt -> {
			if ("lookAndFeel".equals(evt.getPropertyName())) {
				onLafChanged();
			}
		});
	}

	private void onLafChanged() {
		previewBackgroundColor = UIManager.getColor("List.background");
		previewTextColor = UIManager.getColor("TextField.foreground");
		messageColor = UIManager.getColor("Label.foreground");
		checkeredPaint = null;
		SwingUtilities.updateComponentTreeUI(messageLabel);
		if (currPreview != null && currPreview.hasVisualsInterface) {
			showPreview(currPreview.info.file);
		} else {
			repaint();
		}
	}

	private void initializeHandler(Preview preview) {
		boolean isInitialized = false;
		Unknown unknown = preview.handler;
		String file = preview.info.file.toString();

		// Initialize with stream
		if (!isInitialized) {
			PointerByReference initStreamPtr = new PointerByReference();
			HRESULT hr = unknown.QueryInterface(new REFIID(IID_IInitializeWithStream), initStreamPtr);
			if (COMUtils.SUCCEEDED(hr)) {
				InitializeWithStream initStream = new InitializeWithStream(initStreamPtr.getValue());
				PointerByReference streamPtr = new PointerByReference();
				hr = Shlwapi.INSTANCE.SHCreateStreamOnFileEx(new WString(file), Shlwapi.STGM_READ, 0, false, null, streamPtr);
				COMUtils.checkRC(hr);
				Unknown stream = new Unknown(streamPtr.getValue());
				hr = initStream.Initialize(streamPtr.getValue(), Shlwapi.STGM_READ);
				COMUtils.checkRC(hr);
				isInitialized = true;
				stream.Release();
				initStream.Release();
				preview.info.initializerType = "IInitializeWithStream";
			}
		}

		// Initialize with file
		if (!isInitialized) {
			PointerByReference initFilePtr = new PointerByReference();
			HRESULT hr = unknown.QueryInterface(new REFIID(IID_IInitializeWithFile), initFilePtr);
			if (COMUtils.SUCCEEDED(hr)) {
				InitializeWithFile initFile = new InitializeWithFile(initFilePtr.getValue());
				hr = initFile.Initialize(new WString(file), Shlwapi.STGM_READ);
				COMUtils.checkRC(hr);
				isInitialized = true;
				initFile.Release();
				preview.info.initializerType = "IInitializeWithFile";
			}
		}

		// Initialize with shell item
		if (!isInitialized) {
			PointerByReference initItemPtr = new PointerByReference();
			HRESULT hr = unknown.QueryInterface(new REFIID(IID_IInitializeWithItem), initItemPtr);
			if (COMUtils.SUCCEEDED(hr)) {
				// Create IShellItem
				PointerByReference itemPtr = new PointerByReference();
				REFIID riid = new REFIID(IID_IShellItem.getPointer());
				hr = Shell32.INSTANCE.SHCreateItemFromParsingName(new WString(file), null, riid, itemPtr);
				COMUtils.checkRC(hr);
				Unknown shellItem = new Unknown(itemPtr.getValue());
				InitializeWithItem initItem = new InitializeWithItem(initItemPtr.getValue());
				hr = initItem.Initialize(shellItem.getPointer(), Shlwapi.STGM_READ);
				COMUtils.checkRC(hr);
				isInitialized = true;
				shellItem.Release();
				initItem.Release();
				preview.info.initializerType = "IInitializeWithItem";
			}
		}

		if (!isInitialized)
			throw new RuntimeException("Failed to initialize preview handler");
	}

	private void updatePreviewRect() {
		Preview preview = currPreview;
		if (preview == null)
			return;
		if (preview.isPreviewMode && canvasHwnd != null) {
			preview.thread.submit(() -> {
				RECT rect = new RECT();
				User32.INSTANCE.GetClientRect(canvasHwnd, rect);
				((PreviewHandler) preview.handler).SetRect(rect);
			});
		} else if (preview.isThumbnailMode && preview.thumbnailImage != null) {
			// Check if we need to request a bigger thumbnail
			Rectangle b = getBounds();
			b.width = Math.max(0, b.width - thumbnailInsets.left - thumbnailInsets.right);
			b.height = Math.max(0, b.height - thumbnailInsets.top - thumbnailInsets.bottom);
			if (preview.thumbnailImage.getWidth() < b.width && !preview.isBiggestThumbnail) {
				preview.thread.submit(() -> loadPreview(preview));
			}
			repaint();
		}
	}

	void repaintImmediately() {
		paint(getGraphics());
	}

	@Override
	public void paint(Graphics g) {
		BufferStrategy bs = getBufferStrategy();
		if (bs == null) {
			createBufferStrategy(2);
			bs = getBufferStrategy();
		}
		do {
			do {
				Graphics2D g2d = null;
				try {
					g2d = (Graphics2D) bs.getDrawGraphics();

					// Clear background
					g2d.setColor(previewBackgroundColor);
					g2d.fillRect(0, 0, getWidth(), getHeight());

					// Determine available bounds
					Rectangle b = getBounds();
					b.x += thumbnailInsets.left;
					b.y += thumbnailInsets.top;
					b.width = Math.max(0, b.width - thumbnailInsets.left - thumbnailInsets.right);
					b.height = Math.max(0, b.height - thumbnailInsets.top - thumbnailInsets.bottom);

					// Paint thumbnail or message
					Preview preview = currPreview;
					if (preview != null && preview.isPreviewMode) {
						// Painting done by IPreviewHandler
					} else if (preview != null && preview.isThumbnailMode && preview.thumbnailImage != null) {

						// Calculate scaled size
						int w = preview.thumbnailImage.getWidth();
						int h = preview.thumbnailImage.getHeight();
						double scale = Math.min((double) b.width / w, (double) b.height / h);
						int scaledW = w;
						int scaledH = h;
						if (scale < 1) {
							scaledW = (int) Math.round(w * scale);
							scaledH = (int) Math.round(h * scale);
						} else {
							scale = 1;
						}
						int x = thumbnailInsets.left + (b.width - scaledW) / 2;
						int y = thumbnailInsets.top + (b.height - scaledH) / 2;

						// Paint checkered background
						if (showCheckeredBackground && isMouseOver) {
							if (checkeredPaint == null)
								checkeredPaint = createCheckeredPaint();
							g2d.setPaint(checkeredPaint);
							g2d.fillRect(x, y, scaledW, scaledH);
						}

						// Paint thumbnail
						g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
						g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
						g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
						g2d.drawImage(preview.thumbnailImage, x, y, scaledW, scaledH, null);

					} else {
						Font font = messageLabel.getFont();
						StringBuilder fontStyle = new StringBuilder("font-family:'" + font.getName() + "';font-size:" + font.getSize() + "pt;");
						if (font.isBold())
							fontStyle.append("font-weight:bold;");
						if (font.isItalic())
							fontStyle.append("font-style:italic;");
						messageLabel.setText("<html><div style=\"text-align:center;" + fontStyle + "\">" + previewMessage + "</div></html>");
						messageLabel.setOpaque(false);
						messageLabel.setBounds(b);
						messageLabel.setHorizontalAlignment(SwingConstants.CENTER);
						messageLabel.setForeground(messageColor);
						g2d.translate(thumbnailInsets.left, thumbnailInsets.top);
						messageLabel.paint(g2d);
					}
				} finally {
					if (g2d != null)
						g2d.dispose();
				}
			} while (bs.contentsRestored());

			bs.show();
		} while (bs.contentsLost());
	}

	private TexturePaint createCheckeredPaint() {
		boolean isDarkTheme = isDarkTheme(previewBackgroundColor);
		int size = checkeredTileSize;
		BufferedImage texture = new BufferedImage(2 * size, 2 * size, BufferedImage.TYPE_INT_ARGB);
		Graphics2D g2d = texture.createGraphics();

		// Determine colors and make sure the contrast is good enough
		Color color1 = isDarkTheme ? Color.BLACK : Color.WHITE;
		Color color2 = previewBackgroundColor;
		if (isDarkTheme) {
			float[] hsb = getHSB(previewBackgroundColor);
			if (hsb[2] < 0.25f) {
				color2 = Color.getHSBColor(0, 0, 0.25f);
			}
		} else {
			float[] hsb = getHSB(previewBackgroundColor);
			if (hsb[2] > 0.93f) {
				color2 = Color.getHSBColor(hsb[0], hsb[1], 0.93f);
			}
		}

		g2d.setColor(color1);
		g2d.fillRect(0, 0, size, size);
		g2d.fillRect(size, size, size, size);
		g2d.setColor(color2);
		g2d.fillRect(size, 0, size, size);
		g2d.fillRect(0, size, size, size);
		g2d.dispose();
		return new TexturePaint(texture, new Rectangle(0, 0, 2 * size, 2 * size));
	}

	@Override
	public void update(Graphics g) {
		paint(g);
	}

	private BufferedImage hbitmapToBufferedImage(HBITMAP hbmp, int alphaType) {
		BITMAP bm = new BITMAP();
		GDI32.INSTANCE.GetObject(hbmp, bm.size(), bm.getPointer());
		bm.read();
		int width = bm.bmWidth.intValue();
		int height = bm.bmHeight.intValue();

		HDC hdc = GDI32.INSTANCE.CreateCompatibleDC(null);
		HANDLE oldObj = GDI32.INSTANCE.SelectObject(hdc, hbmp);

		BITMAPINFO bmi = new BITMAPINFO();
		bmi.bmiHeader.biSize = bmi.bmiHeader.size();
		bmi.bmiHeader.biWidth = width;
		bmi.bmiHeader.biHeight = -height;
		bmi.bmiHeader.biPlanes = 1;
		bmi.bmiHeader.biBitCount = 32;
		bmi.bmiHeader.biCompression = WinGDI.BI_RGB;

		Memory bits = new Memory((long) width * height * 4);
		GDI32.INSTANCE.GetDIBits(hdc, hbmp, 0, height, bits, bmi, WinGDI.DIB_RGB_COLORS);

		GDI32.INSTANCE.SelectObject(hdc, oldObj);
		GDI32.INSTANCE.DeleteDC(hdc);

		int[] pixels = bits.getIntArray(0, width * height);
		if (alphaType == WTSAT_UNKNOWN || alphaType == WTSAT_RGB) {
			for (int i = 0; i < pixels.length; i++) {
				pixels[i] |= 0xFF000000;
			}
		}

		BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
		image.setRGB(0, 0, width, height, pixels, 0, width);
		return image;
	}

	/**
	 * Opens the given file using the default application in the Windows registry.
	 * 
	 * @param parent used to display any messageboxes that the system might produce while executing this function
	 * @param file the file to open
	 */
	public static void openFile(Component parent, File file) {
		if (file == null)
			return;
		SHELLEXECUTEINFO info = new SHELLEXECUTEINFO();
		info.hwnd = parent == null ? null : new HWND(Native.getComponentPointer(parent));
		info.lpFile = file.toString();
		info.nShow = Shell32.SW_SHOWDEFAULT;
		Shell32.INSTANCE.ShellExecuteEx(info);
	}

	private static boolean isDarkTheme(Color bg) {
		if (bg == null)
			return false;
		double luminance = 0.299 * bg.getRed() + 0.587 * bg.getGreen() + 0.114 * bg.getBlue();
		return luminance < 128;
	}

	private static int clamp(int value, int min, int max) {
		return value < min ? min : value > max ? max : value;
	}

	private static float[] getHSB(Color c) {
		return Color.RGBtoHSB(c.getRed(), c.getGreen(), c.getBlue(), null);
	}

}
