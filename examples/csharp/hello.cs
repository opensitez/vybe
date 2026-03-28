// Simple C# Hello World
// Run: vybec examples/csharp/hello.cs

using System;

Console.WriteLine("Hello from C#!");
Console.WriteLine("2 + 3 = " + (2 + 3));

// Variables
var name = "Vybe";
var version = 1.0;
Console.WriteLine("Welcome to " + name + " v" + version);

// Control flow
for (int i = 1; i <= 5; i++)
{
    Console.WriteLine("Count: " + i);
}

// Class
class Greeter
{
    private string greeting;

    public Greeter(string g)
    {
        this.greeting = g;
    }

    public void Greet(string name)
    {
        Console.WriteLine(greeting + ", " + name + "!");
    }
}

var greeter = new Greeter("Hello");
greeter.Greet("World");
greeter.Greet("C#");
