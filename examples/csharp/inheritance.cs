// C# Inheritance demo — shows class hierarchy, virtual methods, overrides
// Run: vybec examples/csharp/inheritance.cs

using System;

class Shape
{
    public string Name;

    public Shape(string name)
    {
        this.Name = name;
    }

    public virtual double Area()
    {
        return 0;
    }

    public void Describe()
    {
        Console.WriteLine($"{Name}: area = {Area()}");
    }
}

class Circle : Shape
{
    public double Radius;

    public Circle(double r) : base("Circle")
    {
        this.Radius = r;
    }

    public override double Area()
    {
        return 3.14159 * Radius * Radius;
    }
}

class Rectangle : Shape
{
    public double Width;
    public double Height;

    public Rectangle(double w, double h) : base("Rectangle")
    {
        this.Width = w;
        this.Height = h;
    }

    public override double Area()
    {
        return Width * Height;
    }
}

class Square : Rectangle
{
    public Square(double side) : base(side, side)
    {
        this.Name = "Square";
    }
}

// Entry point
var shapes = new Shape[]
{
    new Circle(5),
    new Rectangle(4, 6),
    new Square(3)
};

foreach (var shape in shapes)
{
    shape.Describe();
}

Console.WriteLine("Total shapes: " + shapes.Length);
