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

import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.io.File;
import java.util.List;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.DefaultListCellRenderer;
import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.TransferHandler;
import javax.swing.UIManager;
import javax.swing.UIManager.LookAndFeelInfo;
import javax.swing.border.CompoundBorder;
import javax.swing.border.EmptyBorder;

import stm.ui.inspector.Inspector;
import stm.ui.win32.WindowsPreviewCanvas.PreviewInfo;
import stm.ui.win32.WindowsPreviewCanvas.PreviewListener;

public class WindowsPreviewDemo {

	public static class GridBagConstraintsBuilder {

		GridBagConstraints gbc = new GridBagConstraints();

		public GridBagConstraintsBuilder reset() {
			gbc.gridx = GridBagConstraints.RELATIVE;
			gbc.gridy = GridBagConstraints.RELATIVE;
			gbc.gridwidth = 1;
			gbc.gridheight = 1;
			gbc.weightx = 0;
			gbc.weighty = 0;
			gbc.anchor = GridBagConstraints.WEST;
			gbc.fill = GridBagConstraints.NONE;
			gbc.insets = new Insets(5, 5, 5, 5);
			gbc.ipadx = 0;
			gbc.ipady = 0;
			return this;
		}

		public GridBagConstraintsBuilder xy(int x, int y) {
			reset();
			gbc.gridx = x;
			gbc.gridy = y;
			return this;
		}

		public GridBagConstraintsBuilder span(int spanX, int spanY) {
			gbc.gridwidth = spanX;
			gbc.gridheight = spanY;
			return this;
		}

		public GridBagConstraintsBuilder spanX(int spanX) {
			gbc.gridwidth = spanX;
			return this;
		}

		public GridBagConstraintsBuilder insets(int top, int left, int bottom, int right) {
			gbc.insets = new Insets(top, left, bottom, right);
			return this;
		}

		public GridBagConstraintsBuilder fill() {
			gbc.fill = GridBagConstraints.BOTH;
			return this;
		}

		public GridBagConstraintsBuilder grabX() {
			gbc.weightx = 1;
			return this;
		}

		public GridBagConstraintsBuilder grabY() {
			gbc.weighty = 1;
			return this;
		}

		public GridBagConstraints get() {
			return gbc;
		}

	}

	private static final String FLATLAF_DARK = "com.formdev.flatlaf.FlatDarkLaf";
	private static final String FLATLAF_LIGHT = "com.formdev.flatlaf.FlatLightLaf";

	private JFrame frame;
	private DefaultListModel<File> listModel;
	private JComboBox<LookAndFeelInfo> lafComboBox;
	private JLabel dropLabel;
	private JList<File> fileList;
	private JPanel previewPanel;
	private JSplitPane splitPane;
	private JTextField clsidField;
	private JTextField initalizerField;
	private JTextField interfaceField;
	private WindowsPreviewCanvas previewCanvas;

	public WindowsPreviewDemo() {
		Inspector.enable();
		createUI();

		File dir = new File("F:\\Projekte\\WindowsFilePreviewJNA\\testfiles");
		File[] files = dir.listFiles();
		for (File file : files) {
			if (file.isFile()) {
				listModel.addElement(file);
			}
		}
	}

	private void createUI() {
		try {
			if (isLookAndFeelAvailable(FLATLAF_LIGHT))
				UIManager.installLookAndFeel("FlatLaf Light", FLATLAF_LIGHT);
			if (isLookAndFeelAvailable(FLATLAF_DARK))
				UIManager.installLookAndFeel("FlatLaf Dark", FLATLAF_DARK);
			if (isLookAndFeelAvailable(FLATLAF_LIGHT)) {
				UIManager.setLookAndFeel(FLATLAF_LIGHT);
			} else {
				UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
			}
			UIManager.put("TitlePane.useWindowDecorations", false);
			UIManager.put("ScrollBar.showButtons", true);
			UIManager.put("ScrollBar.width", 16);
		} catch (Exception e) {
			e.printStackTrace();
		}

		frame = new JFrame();
		frame.setTitle("Windows File Preview Demo");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(825, 600);
		frame.setLocationRelativeTo(null);

		splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT);
		splitPane.setBorder(null);
		splitPane.setDividerLocation(400);
		splitPane.setContinuousLayout(true);

		// Left panel
		JPanel leftPanel = new JPanel(new GridBagLayout());
		leftPanel.setBorder(new EmptyBorder(5, 5, 5, 0));
		GridBagConstraintsBuilder gbc = new GridBagConstraintsBuilder();

		// Buttons
		{
			JButton openFileButton = new JButton("Open File");
			openFileButton.addActionListener(e -> {
				JFileChooser fileChooser = new JFileChooser();
				if (fileChooser.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
					listModel.addElement(fileChooser.getSelectedFile());
				}
			});

			JButton openFolderButton = new JButton("Open Folder");
			openFolderButton.addActionListener(e -> {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				if (fileChooser.showOpenDialog(frame) == JFileChooser.APPROVE_OPTION) {
					File dir = fileChooser.getSelectedFile();
					File[] files = dir.listFiles();
					DefaultListModel<File> listModel = new DefaultListModel<>();
					for (File file : files) {
						if (file.isFile()) {
							listModel.addElement(file);
						}
					}
					fileList.setModel(listModel);
				}
			});

			LookAndFeelInfo[] lafs = UIManager.getInstalledLookAndFeels();
			lafComboBox = new JComboBox<>(lafs);
			lafComboBox.setSelectedItem(findCurrentLaf(lafs));
			lafComboBox.addActionListener(e -> setLookAndFeel((LookAndFeelInfo) lafComboBox.getSelectedItem()));
			lafComboBox.setRenderer(new DefaultListCellRenderer() {
				@Override
				public Component getListCellRendererComponent(JList< ? > list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
					LookAndFeelInfo info = (LookAndFeelInfo) value;
					return super.getListCellRendererComponent(list, info.getName(), index, isSelected, cellHasFocus);
				}
			});

			JPanel panel = new JPanel();
			panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
			panel.add(openFileButton);
			panel.add(Box.createHorizontalStrut(10));
			panel.add(openFolderButton);
			panel.add(Box.createHorizontalGlue());
			panel.add(lafComboBox);
			leftPanel.add(panel, gbc.xy(0, 0).spanX(2).fill().get());
		}

		// Drop Area
		{
			JPanel dropPanel = new JPanel(new BorderLayout());
			dropPanel.setBorder(BorderFactory.createTitledBorder(""));
			dropPanel.setPreferredSize(new Dimension(300, 40));
			dropLabel = new JLabel("Drop files here", SwingConstants.CENTER);
			dropPanel.add(dropLabel, BorderLayout.CENTER);

			// Enable drag and drop
			dropPanel.setTransferHandler(new TransferHandler() {
				@Override
				public boolean canImport(TransferSupport support) {
					return support.isDataFlavorSupported(DataFlavor.javaFileListFlavor);
				}
				@Override
				public boolean importData(TransferSupport support) {
					if (!canImport(support))
						return false;
					Transferable transferable = support.getTransferable();
					try {
						@SuppressWarnings("unchecked")
						List<File> files = (List<File>) transferable.getTransferData(DataFlavor.javaFileListFlavor);
						for (File file : files) {
							listModel.addElement(file);
						}
						dropLabel.setText("Dropped successfully!");
						return true;
					} catch (Exception ex) {
						ex.printStackTrace();
						return false;
					}
				}
			});
			leftPanel.add(dropPanel, gbc.xy(0, 1).spanX(2).fill().get());
		}

		// File list
		{
			listModel = new DefaultListModel<>();
			fileList = new JList<>(listModel);
			fileList.getSelectionModel().setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
			fileList.setCellRenderer(new DefaultListCellRenderer() {
				@Override
				public Component getListCellRendererComponent(JList< ? > list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
					setBorder(null);
					super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
					setBorder(new CompoundBorder(getBorder(), new EmptyBorder(2, 2, 2, 2)));
					if (value instanceof File) {
						setText(((File) value).getName());
					}
					return this;
				}
			});
			fileList.addListSelectionListener(e -> {
				if (!e.getValueIsAdjusting()) {
					File selectedFile = fileList.getSelectedValue();
					if (selectedFile != null && fileList.getSelectedIndices().length == 1) {
						previewCanvas.showPreview(selectedFile);
					} else {
						interfaceField.setText("");
						clsidField.setText("");
						initalizerField.setText("");
						previewCanvas.showPreview(null);
					}
				}
			});

			JScrollPane scrollPane = new JScrollPane(fileList);
			scrollPane.setMinimumSize(new Dimension(0, 0));
			scrollPane.setPreferredSize(new Dimension(100, 100)); // prevent it from grabbing the vertical space
			leftPanel.add(scrollPane, gbc.xy(0, 2).spanX(2).grabY().fill().get());
		}

		// Detail fields
		{
			// Interface
			JLabel interfaceLabel = new JLabel("Interface:");
			leftPanel.add(interfaceLabel, gbc.xy(0, 3).get());
			interfaceField = newTextField();
			leftPanel.add(interfaceField, gbc.xy(1, 3).grabX().fill().get());

			// CLSID
			JLabel clsidLabel = new JLabel("CLSID:");
			leftPanel.add(clsidLabel, gbc.xy(0, 4).fill().get());
			clsidField = newTextField();
			leftPanel.add(clsidField, gbc.xy(1, 4).grabX().fill().get());

			// Initializer
			JLabel sizeLabel = new JLabel("Initializer:  ");
			leftPanel.add(sizeLabel, gbc.xy(0, 5).fill().get());
			initalizerField = newTextField();
			leftPanel.add(initalizerField, gbc.xy(1, 5).grabX().fill().get());
		}

		// Right Panel
		previewPanel = new JPanel(new BorderLayout());
		previewCanvas = new WindowsPreviewCanvas();
		previewCanvas.setThumbnailInsets(new Insets(10, 10, 10, 10));
		previewCanvas.addPreviewListener(new PreviewListener() {
			@Override
			public void onPreviewLoaded(PreviewInfo info) {
				interfaceField.setText(info.interfaceType);
				clsidField.setText(info.clsid);
				initalizerField.setText(info.initializerType);
			}
		});
		previewPanel.add(previewCanvas, BorderLayout.CENTER);
		updatePreviewPanelLaf();

		splitPane.setLeftComponent(leftPanel);
		splitPane.setRightComponent(previewPanel);

		frame.add(splitPane);
	}

	private LookAndFeelInfo findCurrentLaf(LookAndFeelInfo[] lafs) {
		String className = UIManager.getLookAndFeel().getClass().getName();
		for (LookAndFeelInfo info : lafs) {
			if (info.getClassName().equals(className))
				return info;
		}
		return null;
	}

	private void updatePreviewPanelLaf() {
		previewPanel.setBorder(new CompoundBorder(new EmptyBorder(10, 5, 10, 10), new JScrollPane().getBorder()));
		previewCanvas.setPreviewFont((Font) UIManager.getLookAndFeelDefaults().get("defaultFont"));
	}

	private void setLookAndFeel(LookAndFeelInfo info) {
		try {
			UIManager.setLookAndFeel(info.getClassName());
			splitPane.setBorder(null);
			updatePreviewPanelLaf();
			SwingUtilities.updateComponentTreeUI(frame);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private JTextField newTextField() {
		JTextField tf = new JTextField();
		tf.setEditable(false);
		return tf;
	}

	private static boolean isLookAndFeelAvailable(String className) {
		try {
			Class< ? > lookAndFeelClass = Class.forName(className);
			return javax.swing.LookAndFeel.class.isAssignableFrom(lookAndFeelClass);
		} catch (ClassNotFoundException e) {
			return false;
		}
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(() -> new WindowsPreviewDemo().frame.setVisible(true));
	}
}