using System;
using System.Drawing;
using System.Windows.Forms;

public partial class MainForm
{
    private Button launchButton;
    private Button boostButton;
    private Button mysteryButton;
    private Button resetButton;
    private Label titleLabel;
    private Label statusLabel;
    private Label scoreLabel;
    private GroupBox panelBox;

    private void InitializeComponent()
    {
        this.launchButton = new Button();
        this.boostButton = new Button();
        this.mysteryButton = new Button();
        this.resetButton = new Button();
        this.titleLabel = new Label();
        this.statusLabel = new Label();
        this.scoreLabel = new Label();
        this.panelBox = new GroupBox();

        this.SuspendLayout();

        this.Text = "Cosmic Arcade";
        this.Name = "MainForm";
        this.ClientSize = new Size(720, 420);
        this.BackColor = Color.Black;

        this.titleLabel.Name = "titleLabel";
        this.titleLabel.Location = new Point(24, 18);
        this.titleLabel.Size = new Size(520, 34);
        this.titleLabel.Text = "Cosmic Arcade: Mission Control";

        this.panelBox.Name = "panelBox";
        this.panelBox.Location = new Point(24, 60);
        this.panelBox.Size = new Size(670, 250);
        this.panelBox.Text = "Control Deck";

        this.launchButton.Name = "launchButton";
        this.launchButton.Location = new Point(45, 110);
        this.launchButton.Size = new Size(130, 44);
        this.launchButton.Text = "Launch";
        this.launchButton.Click += this.launchButton_Click;

        this.boostButton.Name = "boostButton";
        this.boostButton.Location = new Point(190, 110);
        this.boostButton.Size = new Size(130, 44);
        this.boostButton.Text = "Boost";
        this.boostButton.Click += this.boostButton_Click;

        this.mysteryButton.Name = "mysteryButton";
        this.mysteryButton.Location = new Point(335, 110);
        this.mysteryButton.Size = new Size(130, 44);
        this.mysteryButton.Text = "Mystery Crate";
        this.mysteryButton.Click += this.mysteryButton_Click;

        this.resetButton.Name = "resetButton";
        this.resetButton.Location = new Point(480, 110);
        this.resetButton.Size = new Size(130, 44);
        this.resetButton.Text = "Reset";
        this.resetButton.Click += this.resetButton_Click;

        this.statusLabel.Name = "statusLabel";
        this.statusLabel.Location = new Point(24, 330);
        this.statusLabel.Size = new Size(670, 24);
        this.statusLabel.Text = "Status";

        this.scoreLabel.Name = "scoreLabel";
        this.scoreLabel.Location = new Point(24, 360);
        this.scoreLabel.Size = new Size(220, 24);
        this.scoreLabel.Text = "Score: 0";

        this.Controls.Add(this.titleLabel);
        this.Controls.Add(this.panelBox);
        this.Controls.Add(this.launchButton);
        this.Controls.Add(this.boostButton);
        this.Controls.Add(this.mysteryButton);
        this.Controls.Add(this.resetButton);
        this.Controls.Add(this.statusLabel);
        this.Controls.Add(this.scoreLabel);

        this.ResumeLayout(false);
    }
}
