use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(
    auto_property_initializer_sets_default_title,
    r#"
class Article {
    public string Title { get; set; } = "draft";
}
var article = new Article();
Console.WriteLine(article.Title);
article.Title = "published";
Console.WriteLine(article.Title);
"#,
    ["draft", "published"]
);

csharp_case!(
    getter_only_computed_property_reads_backing_fields,
    r#"
class Rectangle {
    public int Width { get; set; }
    public int Height { get; set; }
    public int Area { get { return Width * Height; } }
}
var rectangle = new Rectangle { Width = 4, Height = 6 };
Console.WriteLine(rectangle.Area);
"#,
    ["24"]
);

csharp_case!(
    private_setter_property_changes_through_instance_method,
    r#"
class Counter {
    public int Value { get; private set; }
    public void Increment() { Value++; }
}
var counter = new Counter();
counter.Increment();
counter.Increment();
Console.WriteLine(counter.Value);
"#,
    ["2"]
);

csharp_case!(
    init_only_property_is_set_by_object_initializer,
    r#"
class Customer {
    public string Name { get; init; }
    public int Tier { get; init; }
}
var customer = new Customer { Name = "Ada", Tier = 3 };
Console.WriteLine(customer.Name);
Console.WriteLine(customer.Tier);
"#,
    ["Ada", "3"]
);

csharp_case!(
    expression_bodied_getter_returns_formatted_code,
    r#"
class Package {
    public string Prefix { get; set; }
    public int Number { get; set; }
    public string Code => Prefix + "-" + Number;
}
var package = new Package { Prefix = "PKG", Number = 42 };
Console.WriteLine(package.Code);
"#,
    ["PKG-42"]
);

csharp_case!(
    validated_setter_rejects_negative_values,
    r#"
class Thermometer {
    int celsius;
    public int Celsius {
        get { return celsius; }
        set { celsius = value < 0 ? 0 : value; }
    }
}
var thermometer = new Thermometer();
thermometer.Celsius = -7;
Console.WriteLine(thermometer.Celsius);
thermometer.Celsius = 18;
Console.WriteLine(thermometer.Celsius);
"#,
    ["0", "18"]
);

csharp_case!(
    object_initializer_populates_nested_property_graph,
    r#"
class Address {
    public string City { get; set; }
}
class Office {
    public string Name { get; set; }
    public Address Location { get; set; }
}
var office = new Office {
    Name = "HQ",
    Location = new Address { City = "Paris" }
};
Console.WriteLine(office.Name);
Console.WriteLine(office.Location.City);
"#,
    ["HQ", "Paris"]
);

csharp_case!(
    read_only_property_is_assigned_from_constructor,
    r#"
class BuildInfo {
    public string Version { get; }
    public BuildInfo(string version) { Version = version; }
}
var info = new BuildInfo("1.2.3");
Console.WriteLine(info.Version);
"#,
    ["1.2.3"]
);

csharp_case!(
    expression_bodied_setter_updates_hidden_field,
    r#"
class ScoreBoard {
    int score;
    public int Score {
        get => score;
        set => score = value * 10;
    }
}
var board = new ScoreBoard();
board.Score = 7;
Console.WriteLine(board.Score);
"#,
    ["70"]
);

csharp_case!(
    property_access_uses_base_virtual_getter_override,
    r#"
class BasePerson {
    public virtual string Label { get { return "base"; } }
}
class Employee : BasePerson {
    public override string Label { get { return "employee"; } }
}
BasePerson person = new Employee();
Console.WriteLine(person.Label);
"#,
    ["employee"]
);
