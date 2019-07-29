namespace DRP.Find_Changeset_By_Comment
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;
    using EnvDTE80;
    using Microsoft.TeamFoundation.VersionControl.Client;
    using Microsoft.VisualStudio.Shell.Interop;
    using Microsoft.VisualStudio.TeamFoundation.VersionControl;

    public partial class FindForm : Form
    {
        /// <summary>
        /// The search criteria
        /// </summary>
        private SearchCriteria searchCriteria = new SearchCriteria();

        /// <summary>
        /// The version control server which will be used to find history
        /// </summary>
        private VersionControlServer versionControlServer = null;

        /// <summary>
        /// The version control ext
        /// </summary>
        private VersionControlExt versionControlExt = null;

        /// <summary>
        /// The list of changesets returned from the search
        /// </summary>
        private List<Changeset> changes = new List<Changeset>();

        /// <summary>
        /// The list of affected files
        /// </summary>
        private List<string> files = new List<string>();

        /// <summary>
        /// The UI shell interface
        /// </summary>
        public IVsUIShell UiShell = null;

        /// <summary>
        /// The main automation object
        /// </summary>
        public DTE2 AutomationObject = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="FindForm"/> class.
        /// </summary>
        public FindForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FindForm"/> class.
        /// </summary>
        /// <param name="automationObject">The automation object.</param>
        public FindForm(DTE2 automationObject) : this()
        {
            AutomationObject = automationObject;
            versionControlExt = (VersionControlExt)this.AutomationObject.GetObject("Microsoft.VisualStudio.TeamFoundation.VersionControl.VersionControlExt");
            versionControlServer = versionControlExt.Explorer.Workspace.VersionControlServer;
            ContainingFileText.Text = versionControlExt.Explorer.CurrentFolderItem.SourceServerPath;
            FromDatePicker.Value = DateTime.Now - new TimeSpan(28, 0, 0, 0);
            ToDatePicker.Value = DateTime.Now;
        }

        /// <summary>
        /// Handles the Click event of the FindButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void FindButton_Click(object sender, EventArgs e)
        {
            // Check that we have a folder to search
            if (string.IsNullOrWhiteSpace(ContainingFileText.Text))
            {
                MessageBox.Show("Please specify a source control folder to search", "Find Changeset By Comment", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // Check that we have a valid regular expression
            if (useRegularExpressions.Checked && !string.IsNullOrWhiteSpace(ContainingCommentText.Text))
            {
                try
                {
                    Regex rx = new Regex(ContainingCommentText.Text);
                }
                catch (ArgumentException)
                {
                    MessageBox.Show("Please specify a valid regular expression", "Find Changeset By Comment", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            searchCriteria.SearchFile = ContainingFileText.Text.Trim();
            searchCriteria.SearchUser = ByUserText.Text;
            searchCriteria.FromDateVersion = FromDatePicker.Checked ? new DateVersionSpec(FromDatePicker.Value) : null;
            searchCriteria.ToDateVersion = ToDatePicker.Checked ? new DateVersionSpec(ToDatePicker.Value) : null;
            searchCriteria.SearchComment = ContainingCommentText.Text;
            searchCriteria.UseRegex = useRegularExpressions.Checked;
            ResultsList.Items.Clear();
            FilesList.Items.Clear();
            FindButton.Enabled = false;
            FindChangesetsWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Handles the Click event of the CloseButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void CloseButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Handles the DoubleClick event of the ResultsList control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void ResultsList_DoubleClick(object sender, EventArgs e)
        {
            if (ResultsList.SelectedItems.Count == 0)
            {
                return;
            }

            foreach (Changeset changeset in changes)
            {
                if (changeset.ChangesetId.ToString() == ResultsList.SelectedItems[0].Text)
                {
                    versionControlExt.ViewChangesetDetails(changeset.ChangesetId);
                    break;
                }
            }
        }

        /// <summary>
        /// Handles the Load event of the FindForm control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void FindForm_Load(object sender, EventArgs e)
        {
            // Do any initiailization here
        }

        /// <summary>
        /// Handles the Click event of the CopyButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void CopyButton_Click(object sender, EventArgs e)
        {
            var clipboardText = new StringBuilder();
            if (ChangesetsAndFiles.SelectedTab == ChangesetList)
            {
                // Copy list of changesets
                foreach (Changeset changeset in changes)
                {
                    clipboardText.AppendLine(string.Format("{0}\t{1}\t{2}\t{3}", changeset.ChangesetId, changeset.Owner, changeset.CreationDate, changeset.Comment));
                }
            }
            else
            {
                // Copy list of files
                foreach (string listItem in FilesList.Items)
                {
                    clipboardText.AppendLine(listItem);
                }
            }

            if (clipboardText != null)
            {
                Clipboard.SetText(clipboardText.ToString());
            }
            else
            {
                MessageBox.Show("No results to copy.", "Find Changeset By Comment", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Handles the DoWork event of the FindChangesetsWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="DoWorkEventArgs"/> instance containing the event data.</param>
        private void FindChangesetsWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            // Do not access the form's BackgroundWorker reference directly.
            // Instead, use the reference provided by the sender parameter.
            var bw = sender as BackgroundWorker;

            // Start looking for changesets
            e.Result = FindChangesetsByComment(bw, searchCriteria);

            // If the operation was canceled by the user, 
            // set the DoWorkEventArgs.Cancel property to true.
            if (bw.CancellationPending)
            {
                e.Cancel = true;
            }
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the FindChangesetsWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RunWorkerCompletedEventArgs"/> instance containing the event data.</param>
        private void FindChangesetsWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                // The user canceled the operation.
                MessageBox.Show("Operation was canceled", "Find Changeset By Comment", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (e.Error != null)
            {
                // There was an error during the operation.
                string msg = String.Format("An error occurred: {0}", e.Error.Message);
                MessageBox.Show(msg);
            }
            else
            {
                // The operation completed normally.
                // Check for zero things found
                if (changes.Count == 0)
                {
                    MessageBox.Show("No changesets found matching the search criteria", "Find Changeset By Comment", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

            FindButton.Enabled = true;
        }

        /// <summary>
        /// Finds the changesets by comment.
        /// </summary>
        /// <param name="bw">The background worker.</param>
        /// <param name="sc">The search criteria.</param>
        /// <returns>An error code</returns>
        private int FindChangesetsByComment(BackgroundWorker bw, SearchCriteria sc)
        {
            changes.Clear();
            files.Clear();
            try
            {
                IEnumerable changesets = versionControlServer.QueryHistory(
                    sc.SearchFile,
                    VersionSpec.Latest,
                    0,
                    RecursionType.Full,
                    string.IsNullOrWhiteSpace(sc.SearchUser) ? null : sc.SearchUser,
                    sc.FromDateVersion,
                    sc.ToDateVersion,
                    Int32.MaxValue,
                    true,
                    false);
                foreach (Changeset changeset in changesets)
                {
                    if (bw.CancellationPending)
                    {
                        break;
                    }

                    if (string.IsNullOrWhiteSpace(sc.SearchComment) ||
                        (!sc.UseRegex && changeset.Comment.IndexOf(sc.SearchComment.Trim(), StringComparison.CurrentCultureIgnoreCase) != -1) ||
                        (sc.UseRegex && Regex.IsMatch(changeset.Comment, sc.SearchComment)))
                    {
                        changes.Add(changeset);
                        bw.ReportProgress(0, changeset);
                        foreach (Change change in changeset.Changes)
                        {
                            if (!files.Contains(change.Item.ServerItem))
                            {
                                files.Add(change.Item.ServerItem);
                                bw.ReportProgress(0, change.Item.ServerItem);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Find Changeset By Comment", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return 0;
            }

            return 0;
        }

        /// <summary>
        /// Handles the Click event of the CancelButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void CancelFindButton_Click(object sender, EventArgs e)
        {
            if (FindChangesetsWorker.IsBusy)
            {
                FindChangesetsWorker.CancelAsync();
            }
        }

        /// <summary>
        /// Handles the ProgressChanged event of the FindChangesetsWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="ProgressChangedEventArgs"/> instance containing the event data.</param>
        private void FindChangesetsWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.UserState is Changeset)
            {
                var changeset = e.UserState as Changeset;
                ResultsList.Items.Add(new ListViewItem(new string[] { changeset.ChangesetId.ToString(), changeset.Owner, changeset.CreationDate.ToString(), changeset.Comment }));
            }
            else
            {
                string file = e.UserState as string;
                FilesList.Items.Add(file);
            }
        }

        /// <summary>
        /// The search criteria for finding changesets
        /// </summary>
        private class SearchCriteria
        {
            /// <summary>
            /// Gets or sets the file/folder to search
            /// </summary>
            public string SearchFile
            {
                get;
                set;
            }

            /// <summary>
            /// Gets or sets the user who made the change
            /// </summary>
            public string SearchUser
            {
                get;
                set;
            }

            /// <summary>
            /// Gets or sets the start date
            /// </summary>
            public DateVersionSpec FromDateVersion
            {
                get;
                set;
            }

            /// <summary>
            /// Gets or sets the end date
            /// </summary>
            public DateVersionSpec ToDateVersion
            {
                get;
                set;
            }

            /// <summary>
            /// Gets or sets the comment that we're searching for
            /// </summary>
            public string SearchComment
            {
                get;
                set;
            }

            /// <summary>
            /// Gets or sets a value indicating whether we are using regular expressions
            /// </summary>
            public bool UseRegex
            {
                get;
                set;
            }
        }

        /// <summary>
        /// Handles the KeyUp event of the ResultsList control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="KeyEventArgs"/> instance containing the event data.</param>
        private void ResultsList_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
            {
                ResultsList_DoubleClick(sender, null);
            }
        }
    }
}
