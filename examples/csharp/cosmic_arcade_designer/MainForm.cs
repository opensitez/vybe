using System;
using System.Windows.Forms;

public partial class MainForm : Form
{
    private int score = 0;
    private int boostLevel = 1;

    public MainForm()
    {
        InitializeComponent();
        statusLabel.Text = "Welcome, Captain. Launch for points.";
        scoreLabel.Text = "Score: 0";
    }

    private void launchButton_Click(object sender, EventArgs e)
    {
        score += 10 * boostLevel;
        statusLabel.Text = "Rocket launched. Stardust collected.";
        scoreLabel.Text = "Score: " + score;
    }

    private void boostButton_Click(object sender, EventArgs e)
    {
        if (boostLevel < 5)
        {
            boostLevel += 1;
            statusLabel.Text = "Boost upgraded to x" + boostLevel + ".";
        }
        else
        {
            statusLabel.Text = "Boost is maxed. You are unstoppable.";
        }
    }

    private void mysteryButton_Click(object sender, EventArgs e)
    {
        int reward = (score % 7 + 1) * 3;
        score += reward;
        statusLabel.Text = "Mystery crate opened: +" + reward + " points!";
        scoreLabel.Text = "Score: " + score;
    }

    private void resetButton_Click(object sender, EventArgs e)
    {
        score = 0;
        boostLevel = 1;
        statusLabel.Text = "Mission reset. New galaxy, new chance.";
        scoreLabel.Text = "Score: 0";
    }
}
