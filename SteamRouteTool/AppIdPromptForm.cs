using System;
using System.Globalization;
using System.Windows.Forms;
using SteamRouteTool.Models;

namespace SteamRouteTool
{
    /// <summary>
    /// Asks which Steam app's routes to work on. The list is a convenience only; any app id
    /// the Steam Web API recognises can be typed instead.
    /// </summary>
    public partial class AppIdPromptForm : Form
    {
        public AppIdPromptForm(int initialAppId)
        {
            InitializeComponent();

            foreach (SteamGame game in KnownGames.All)
            {
                cboGame.Items.Add(game);
            }

            SelectAppId(initialAppId);
            cboGame.TextChanged += (sender, e) => HideError();
        }

        /// <summary>The chosen app id. Only meaningful when the dialog returned <see cref="DialogResult.OK"/>.</summary>
        public int AppId { get; private set; }

        private void SelectAppId(int appId)
        {
            foreach (object item in cboGame.Items)
            {
                var game = item as SteamGame;
                if (game != null && game.AppId == appId)
                {
                    cboGame.SelectedItem = game;
                    return;
                }
            }

            // Not a known game, so show the raw id and let the user edit it.
            cboGame.Text = appId > 0
                ? appId.ToString(CultureInfo.InvariantCulture)
                : AppSettings.DefaultAppId.ToString(CultureInfo.InvariantCulture);
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            int appId;
            if (!KnownGames.TryParseAppId(cboGame.Text, out appId))
            {
                ShowError("Enter a Steam app ID, for example 440.");
                return;
            }

            AppId = appId;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void ShowError(string message)
        {
            lblError.Text = message;
            lblError.Visible = true;
            cboGame.Focus();
            cboGame.SelectAll();
        }

        private void HideError()
        {
            if (lblError.Visible) lblError.Visible = false;
        }
    }
}
