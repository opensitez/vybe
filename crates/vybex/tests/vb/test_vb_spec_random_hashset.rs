use super::helpers::run_vb;

macro_rules! vb_expr_spec {
    ($name:ident, $body:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let output = run_vb($body);
            assert_eq!(output, vec![$expected.to_string()]);
        }
    };
}

vb_expr_spec!(
    random_spec_rnd_returns_unit_interval,
    r#"
Module Program
    Sub Main()
        Randomize()
        Dim value As Double = Rnd()
        Console.WriteLine(value >= 0 AndAlso value < 1)
    End Sub
End Module
"#,
    "True"
);

vb_expr_spec!(
    random_spec_system_random_next_max,
    r#"
Imports System
Module Program
    Sub Main()
        Dim rng As New Random()
        Dim value As Integer = rng.Next(10)
        Console.WriteLine(value >= 0 AndAlso value < 10)
    End Sub
End Module
"#,
    "True"
);

vb_expr_spec!(
    random_spec_system_random_next_range,
    r#"
Imports System
Module Program
    Sub Main()
        Dim rng As New Random()
        Dim value As Integer = rng.Next(10, 20)
        Console.WriteLine(value >= 10 AndAlso value < 20)
    End Sub
End Module
"#,
    "True"
);

vb_expr_spec!(
    random_spec_system_random_next_double,
    r#"
Imports System
Module Program
    Sub Main()
        Dim rng As New Random()
        Dim value As Double = rng.NextDouble()
        Console.WriteLine(value >= 0 AndAlso value < 1)
    End Sub
End Module
"#,
    "True"
);

#[test]
fn collections_spec_hashset_add_contains_remove_count() {
    let output = run_vb(
        r#"
Imports System.Collections.Generic
Module Program
    Sub Main()
        Dim items As New HashSet(Of Integer)()
        items.Add(1)
        items.Add(2)
        items.Add(2)
        Console.WriteLine(items.Count)
        Console.WriteLine(items.Contains(2))
        items.Remove(1)
        Console.WriteLine(items.Contains(1))
    End Sub
End Module
"#,
    );

    assert_eq!(output, vec!["2", "True", "False"]);
}
