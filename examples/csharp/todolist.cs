// C# Todo List with WinForms-style GUI
// Run: vybec examples/csharp/todolist.cs

using System;
using System.Windows.Forms;
using System.Drawing;

public class TodoApp : Form
{
    private TextBox txtInput;
    private ListBox lstTodos;
    private Button btnAdd;
    private Button btnRemove;

    public TodoApp()
    {
        this.Text = "Todo List";
        this.ClientSize = new Size(350, 400);

        // Input field
        txtInput = new TextBox();
        txtInput.Name = "txtInput";
        txtInput.Location = new Point(10, 10);
        txtInput.Size = new Size(240, 25);
        this.Controls.Add(txtInput);

        // Add button
        btnAdd = new Button();
        btnAdd.Name = "btnAdd";
        btnAdd.Text = "Add";
        btnAdd.Location = new Point(260, 10);
        btnAdd.Size = new Size(70, 25);
        this.Controls.Add(btnAdd);
        btnAdd.Click += this.OnAddClick;

        // Todo list
        lstTodos = new ListBox();
        lstTodos.Name = "lstTodos";
        lstTodos.Location = new Point(10, 45);
        lstTodos.Size = new Size(320, 300);
        this.Controls.Add(lstTodos);

        // Remove button
        btnRemove = new Button();
        btnRemove.Name = "btnRemove";
        btnRemove.Text = "Remove Selected";
        btnRemove.Location = new Point(10, 355);
        btnRemove.Size = new Size(120, 30);
        this.Controls.Add(btnRemove);
        btnRemove.Click += this.OnRemoveClick;
    }

    private void OnAddClick(object sender, object e)
    {
        AddTodo();
    }

    private void OnRemoveClick(object sender, object e)
    {
        RemoveSelected();
    }

    public void AddTodo()
    {
        var text = txtInput.Text;
        if (text != "")
        {
            lstTodos.Items.Add(text);
            txtInput.Text = "";
        }
    }

    public void RemoveSelected()
    {
        int index = lstTodos.SelectedIndex;
        if (index >= 0)
        {
            lstTodos.Items.RemoveAt(index);
        }
    }
}

var app = new TodoApp();
Application.Run(app);
