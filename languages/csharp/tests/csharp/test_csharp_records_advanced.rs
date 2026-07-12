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
    positional_record_equality_compares_by_value,
    r#"record Point(int X, int Y); Console.WriteLine(new Point(1, 2) == new Point(1, 2));"#,
    ["True"]
);
csharp_case!(
    positional_record_to_string_includes_member_values,
    r#"record Point(int X, int Y); Console.WriteLine(new Point(3, 4).ToString().Contains("X = 3"));"#,
    ["True"]
);
csharp_case!(
    with_expression_copies_record_and_changes_one_member,
    r#"record User(string Name, int Age); var user = new User("Ada", 20); var updated = user with { Age = 21 }; Console.WriteLine(user.Age); Console.WriteLine(updated.Age);"#,
    ["20", "21"]
);
csharp_case!(
    record_deconstruction_returns_positional_members,
    r#"record Point(int X, int Y); var (x, y) = new Point(8, 9); Console.WriteLine(x + y);"#,
    ["17"]
);
csharp_case!(
    record_property_can_be_read_after_construction,
    r#"record Config(string Name); var config = new Config("debug"); Console.WriteLine(config.Name);"#,
    ["debug"]
);
csharp_case!(
    record_with_additional_method_can_compute_value,
    r#"record Counter(int Value) { public int Double() { return Value * 2; } } Console.WriteLine(new Counter(6).Double());"#,
    ["12"]
);
csharp_case!(
    record_inheritance_preserves_base_members,
    r#"record Animal(string Name); record Dog(string Name, int Age) : Animal(Name); var dog = new Dog("Rex", 5); Console.WriteLine(dog.Name); Console.WriteLine(dog.Age);"#,
    ["Rex", "5"]
);
csharp_case!(
    record_with_expression_keeps_untouched_members,
    r#"record Item(string Name, int Count); var item = new Item("pen", 2); var changed = item with { Count = 5 }; Console.WriteLine(changed.Name);"#,
    ["pen"]
);
csharp_case!(
    record_struct_stores_value_semantics,
    r#"record struct Pixel(int X, int Y); var pixel = new Pixel(2, 3); Console.WriteLine(pixel.X + pixel.Y);"#,
    ["5"]
);
csharp_case!(
    readonly_record_struct_exposes_members,
    r#"readonly record struct Size(int Width, int Height); var size = new Size(4, 6); Console.WriteLine(size.Width * size.Height);"#,
    ["24"]
);
csharp_case!(
    record_can_have_init_property_with_default_value,
    r#"record User { public string Name { get; init; } = "guest"; } Console.WriteLine(new User().Name);"#,
    ["guest"]
);
csharp_case!(
    record_clone_via_with_does_not_mutate_original_instance,
    r#"record User(string Name, int Age); var before = new User("Ada", 30); var after = before with { Name = "Grace" }; Console.WriteLine(before.Name); Console.WriteLine(after.Name);"#,
    ["Ada", "Grace"]
);
csharp_case!(
    record_equality_detects_different_member_values,
    r#"record Point(int X, int Y); Console.WriteLine(new Point(1, 2) == new Point(2, 1));"#,
    ["False"]
);
csharp_case!(
    record_hash_code_matches_for_equal_values,
    r#"record Point(int X, int Y); var left = new Point(5, 7); var right = new Point(5, 7); Console.WriteLine(left.GetHashCode() == right.GetHashCode());"#,
    ["True"]
);
csharp_case!(
    record_can_override_to_string_for_custom_format,
    r#"record User(string Name) { public override string ToString() { return $"User:{Name}"; } } Console.WriteLine(new User("Ada"));"#,
    ["User:Ada"]
);
csharp_case!(
    record_with_mutable_property_can_be_updated_after_construction,
    r#"record Box { public int Value { get; set; } } var box = new Box { Value = 3 }; box.Value = 8; Console.WriteLine(box.Value);"#,
    ["8"]
);
csharp_case!(
    record_struct_equality_compares_member_values,
    r#"record struct Pixel(int X, int Y); Console.WriteLine(new Pixel(1, 1) == new Pixel(1, 1));"#,
    ["True"]
);
csharp_case!(
    record_inheritance_to_string_mentions_derived_members,
    r#"record Animal(string Name); record Cat(string Name, string Color) : Animal(Name); Console.WriteLine(new Cat("Milo", "Black").ToString().Contains("Color = Black"));"#,
    ["True"]
);
csharp_case!(
    record_can_define_computed_property_from_primary_members,
    r#"record Rectangle(int Width, int Height) { public int Area => Width * Height; } Console.WriteLine(new Rectangle(3, 7).Area);"#,
    ["21"]
);
csharp_case!(
    record_with_non_positional_members_can_use_object_initializer,
    r#"record Theme { public string Name { get; init; } public int Version { get; init; } } var theme = new Theme { Name = "light", Version = 2 }; Console.WriteLine(theme.Name + ":" + theme.Version);"#,
    ["light:2"]
);
