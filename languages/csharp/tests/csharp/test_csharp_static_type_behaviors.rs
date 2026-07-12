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
    static_field_is_shared_across_instances,
    r#"
class Session {
    public static int Count = 0;
    public Session() { Count++; }
}
new Session();
new Session();
Console.WriteLine(Session.Count);
"#,
    ["2"]
);

csharp_case!(
    static_property_tracks_number_of_created_objects,
    r#"
class Token {
    public static int Created { get; private set; }
    public Token() { Created++; }
}
new Token();
new Token();
new Token();
Console.WriteLine(Token.Created);
"#,
    ["3"]
);

csharp_case!(
    static_constructor_runs_before_first_member_access,
    r#"
class Registry {
    public static string Label;
    static Registry() { Label = "ready"; }
}
Console.WriteLine(Registry.Label);
"#,
    ["ready"]
);

csharp_case!(
    static_readonly_field_exposes_precomputed_value,
    r#"
class Build {
    public static readonly string Channel = "stable";
}
Console.WriteLine(Build.Channel);
"#,
    ["stable"]
);

csharp_case!(
    static_factory_method_returns_initialized_instance,
    r#"
class User {
    public string Name { get; set; }
    public static User CreateAdmin() { return new User { Name = "root" }; }
}
var user = User.CreateAdmin();
Console.WriteLine(user.Name);
"#,
    ["root"]
);

csharp_case!(
    instance_method_calls_static_helper_on_same_type,
    r#"
class Converter {
    public static int Double(int value) { return value * 2; }
    public int Convert(int value) { return Double(value) + 1; }
}
var converter = new Converter();
Console.WriteLine(converter.Convert(5));
"#,
    ["11"]
);

csharp_case!(
    static_dictionary_state_is_mutated_by_each_instance,
    r#"
using System.Collections.Generic;
class Tracker {
    static Dictionary<string, int> counts = new Dictionary<string, int>();
    public void Hit(string key) {
        if (!counts.ContainsKey(key)) counts[key] = 0;
        counts[key]++;
    }
    public static int Read(string key) { return counts[key]; }
}
var a = new Tracker();
var b = new Tracker();
a.Hit("api");
b.Hit("api");
Console.WriteLine(Tracker.Read("api"));
"#,
    ["2"]
);

csharp_case!(
    nested_static_class_exposes_namespaced_helper_method,
    r#"
class TextTools {
    public static class Parts {
        public static string Join(string a, string b) { return a + "/" + b; }
    }
}
Console.WriteLine(TextTools.Parts.Join("a", "b"));
"#,
    ["a/b"]
);

csharp_case!(
    static_field_initializer_uses_expression_result,
    r#"
class Limits {
    public static int Max = 8 * 8;
}
Console.WriteLine(Limits.Max);
"#,
    ["64"]
);

csharp_case!(
    generic_static_field_is_scoped_per_closed_type,
    r#"
class Cache<T> {
    public static int Hits;
}
Cache<int>.Hits++;
Cache<int>.Hits++;
Cache<string>.Hits++;
Console.WriteLine(Cache<int>.Hits);
Console.WriteLine(Cache<string>.Hits);
"#,
    ["2", "1"]
);
