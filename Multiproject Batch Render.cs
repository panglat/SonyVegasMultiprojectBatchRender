/**
 * Sample script that performs batch renders with GUI for selecting
 * render templates.
 *
 * Revision Date: Jun. 28, 2006.
 **/
using System;
using System.IO;
using System.Text;
using System.Drawing;
using System.Collections;
using System.Diagnostics;
using System.Windows.Forms;

using Sony.Vegas;

public class EntryPoint {
	Form dlog;
    Button BrowseButton;
	ComboBox RenderNameComboBox;
	ComboBox RenderTemplateComboBox;
	string[] fileNames = null;
    Button okButton;

    Sony.Vegas.Vegas myVegas = null;

	public void FromVegas(Vegas vegas)
    {
        myVegas = vegas;
        
        dlog = new Form();
        dlog.Text = "Batch Render";
        dlog.MaximizeBox = false;
        dlog.StartPosition = FormStartPosition.CenterScreen;
        dlog.Width = 600;
        int titleBarHeight = dlog.Height - dlog.ClientSize.Height;
        int buttonWidth = 80;
        int buttonTop = dlog.Bottom - 30;
        int buttonsLeft = dlog.Width - (2*(buttonWidth+10));


        BrowseButton = new Button();
        BrowseButton.Left = 4;
        BrowseButton.Top = 4;
        BrowseButton.Width = buttonWidth;
        BrowseButton.Height = BrowseButton.Font.Height + 12;
        BrowseButton.Text = "Browse...";
        BrowseButton.Click += new EventHandler(this.BrowseButton_Click);
        dlog.Controls.Add(BrowseButton);

		
		RenderNameComboBox = new ComboBox();
        RenderNameComboBox.Left = 4;
        RenderNameComboBox.Top = BrowseButton.Bottom + 4;
        RenderNameComboBox.Width = 300;
        RenderNameComboBox.Height = RenderNameComboBox.Font.Height + 12;
		fillRenderNameComboBox();
        RenderNameComboBox.	SelectedIndexChanged += new EventHandler(this.RenderNameComboBox_SelectedIndexChanged);
        dlog.Controls.Add(RenderNameComboBox);


		RenderTemplateComboBox = new ComboBox();
        RenderTemplateComboBox.Left = 4;
        RenderTemplateComboBox.Top = RenderNameComboBox.Bottom + 4;
        RenderTemplateComboBox.Width = 300;
        RenderTemplateComboBox.Height = RenderTemplateComboBox.Font.Height + 12;
		fillRenderTemplateComboBox();
        dlog.Controls.Add(RenderTemplateComboBox);


        okButton = new Button();
        okButton.Text = "OK";
        okButton.Left = dlog.Width - (2*(buttonWidth+10));
        okButton.Top = buttonTop;
        okButton.Width = buttonWidth;
        okButton.Height = okButton.Font.Height + 12;
        okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
        dlog.AcceptButton = okButton;
        okButton.Click += new EventHandler(this.okButton_Click);

        dlog.Controls.Add(okButton);

        Button cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.Left = dlog.Width - (1*(buttonWidth+10));
        cancelButton.Top = buttonTop;
        cancelButton.Height = cancelButton.Font.Height + 12;
        cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        dlog.CancelButton = cancelButton;
        dlog.Controls.Add(cancelButton);

        dlog.Height = titleBarHeight + okButton.Bottom + 8;
        dlog.ShowInTaskbar = false;


		dlog.ShowDialog(myVegas.MainWindow);
	}

	void fillRenderNameComboBox()
	{
		RenderNameComboBox.Items.Clear();
        foreach (Renderer renderer in myVegas.Renderers) {
            try {
                String rendererName = renderer.FileTypeName;
				RenderNameComboBox.Items.Add(rendererName);
            } catch {
                // skip it
            }
        }

		RenderNameComboBox.SelectedIndex = 0;
	}

	void fillRenderTemplateComboBox ()
	{
		Renderer renderer = myVegas.Renderers.FindByName(RenderNameComboBox.SelectedItem.ToString());
		RenderTemplateComboBox.Items.Clear();

		foreach (RenderTemplate renderTemplate in renderer.Templates) {
			try {
				// filter out invalid templates
				if (!renderTemplate.IsValid()) {
					continue;
				}
				// filter out templates that don't have
				// exactly one file extension
				String[] extensions = renderTemplate.FileExtensions;
				if (1 != extensions.Length) {
					continue;
				}
				String templateName = renderTemplate.Name;
				RenderTemplateComboBox.Items.Add(templateName);
			} catch (Exception e) {
				// skip it
				MessageBox.Show(e.ToString());
			}
		}

		RenderTemplateComboBox.SelectedIndex = 0;
	}

    void BrowseButton_Click(Object sender, EventArgs args)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "Sony Vegas Project Files (*.veg)|*.veg";
        openFileDialog.CheckPathExists = true;
        openFileDialog.AddExtension = false;
        openFileDialog.Multiselect  = true;

        if (System.Windows.Forms.DialogResult.OK == openFileDialog.ShowDialog()) {
            this.fileNames = openFileDialog.FileNames;
        } else {
			this.fileNames = null;
		}

    }

    void RenderNameComboBox_SelectedIndexChanged(Object sender, EventArgs args)
    {
		fillRenderTemplateComboBox();
    }

	
    void okButton_Click(Object sender, EventArgs args)
    {
		if (fileNames != null)
		{
			Renderer renderer = myVegas.Renderers.FindByName(RenderNameComboBox.SelectedItem.ToString());
			RenderTemplate renderTemplate = renderer.Templates.FindByName(RenderTemplateComboBox.SelectedItem.ToString());
			
			dlog.Close();
			dlog = null;
			foreach (string fileName in fileNames) {
				String outputFile = Path.ChangeExtension(fileName, renderTemplate.FileExtensions[0].TrimStart('*'));
				myVegas.OpenFile (fileName);
				myVegas.WaitForIdle();
				myVegas.Render(outputFile, renderTemplate);
			}

		}
    }

}