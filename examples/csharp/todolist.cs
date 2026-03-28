// C# Todo List with WinForms-style GUI
// Run: vybec examples/csharp/todolist.cs

using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;

public class TodoApp : Form
{
    private TextBox txtInput;
    private ListBox lstTodos;
    private Button btnAdd;
    private Button btnRemove;
    private List<string> todos = new List<string>();

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
    }

    public void AddTodo()
    {
        var text = txtInput.Text;
        if (text != "")
        {
            todos.Add(text);
            txtInput.Text = "";
            Console.WriteLine("Added: " + text);
        }
    }

    public void RemoveSelected()
    {
        Console.WriteLine("Remove selected todo");
    }
}

var app = new TodoApp();
Application.Run(app);
