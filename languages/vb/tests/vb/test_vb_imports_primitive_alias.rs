use super::helpers::run_vb;

macro_rules! vb_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_case!(
    primitive_alias_integer_math_sum,
    r#"
Imports MyInt = System.Int32

Module M
    Sub Main()
        Dim first As MyInt = 17
        Dim second As MyInt = 25
        Console.WriteLine(first + second)
    End Sub
End Module
"#,
    ["42"]
);

vb_case!(
    primitive_alias_integer_is_reference_boundary,
    r#"
Imports MyInt = System.Int32

Module M
    Sub Main()
        Dim left As MyInt = 11
        Dim right As MyInt = 3
        left = left + (right * MyInt(2))
        Console.WriteLine(left)
    End Sub
End Module
"#,
    ["17"]
);

vb_case!(
    primitive_alias_string_mutation_with_alias,
    r#"
Imports MyText = System.String

Module M
    Sub Main()
        Dim label As MyText = "VB"
        label = label & "-Alias"
        Console.WriteLine(label)
    End Sub
End Module
"#,
    ["VB-Alias"]
);

vb_case!(
    primitive_alias_boolean_logic,
    r#"
Imports MyFlag = System.Boolean

Module M
    Sub Main()
        Dim enabled As MyFlag = True
        Console.WriteLine(Not enabled)
        Console.WriteLine(enabled AndAlso False)
    End Sub
End Module
"#,
    ["False", "False"]
);

vb_case!(
    primitive_alias_char_codepoint,
    r#"
Imports MyCharacter = System.Char

Module M
    Sub Main()
        Dim letter As MyCharacter = "Z"c
        Console.WriteLine(CInt(letter))
    End Sub
End Module
"#,
    ["90"]
);

vb_case!(
    primitive_alias_byte_math_and_convert_back,
    r#"
Imports MyByte = System.Byte

Module M
    Sub Main()
        Dim left As MyByte = 8
        Dim right As MyByte = 5
        Dim total As MyByte = CByte(left + right)
        Console.WriteLine(CInt(total))
    End Sub
End Module
"#,
    ["13"]
);

vb_case!(
    primitive_alias_float_multiplication,
    r#"
Imports MyFloat = System.Single

Module M
    Sub Main()
        Dim a As MyFloat = 1.25
        Dim b As MyFloat = 8.0
        Console.WriteLine(CInt(a * b))
    End Sub
End Module
"#,
    ["10"]
);

vb_case!(
    primitive_alias_double_roundtrip_format,
    r#"
Imports MyFloat = System.Double

Module M
    Sub Main()
        Dim value As MyFloat = 12.5
Console.WriteLine(value.ToString("G"))
        Console.WriteLine(CInt(System.Math.Abs(-17.0)))
    End Sub
End Module
"#,
    ["12.5", "17"]
);

vb_case!(
    primitive_alias_datetime_parts,
    r#"
Imports MyDate = System.DateTime

Module M
    Sub Main()
        Dim stamp As MyDate = MyDate.Parse("2026-07-30T09:30:00")
        Console.WriteLine(stamp.Year)
        Console.WriteLine(stamp.Month)
        Console.WriteLine(stamp.Day)
    End Sub
End Module
"#,
    ["2026", "7", "30"]
);

vb_case!(
    primitive_alias_nullable_int_present,
    r#"
Imports MyNullableInt = System.Nullable(Of Integer)

Module M
    Sub Main()
        Dim value As MyNullableInt = 99
        Dim missing As MyNullableInt = Nothing
        Console.WriteLine(If(value.HasValue, "present", "none"))
        Console.WriteLine(If(missing.HasValue, "present", "none"))
    End Sub
End Module
"#,
    ["present", "none"]
);

vb_case!(
    primitive_alias_nullable_unwrap_value,
    r#"
Imports MyNullableInt = System.Nullable(Of Integer)

Module M
    Sub Main()
        Dim value As MyNullableInt = 42
        Console.WriteLine(CInt(value.Value))
    End Sub
End Module
"#,
    ["42"]
);

vb_case!(
    primitive_alias_array_sum,
    r#"
Imports MyInt = System.Int32

Module M
    Sub Main()
        Dim values() As MyInt = {1, 2, 3, 4}
        Dim total As MyInt = 0
        For Each item As MyInt In values
            total = total + item
        Next
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["10"]
);

vb_case!(
    primitive_alias_list_of_integers,
    r#"
Imports IntList = System.Collections.Generic.List(Of Integer)

Module M
    Sub Main()
        Dim items As New IntList()
        items.Add(4)
        items.Add(5)
        items.Add(1)
        Console.WriteLine(items.Count)
        Console.WriteLine(items(0) + items(1) + items(2))
    End Sub
End Module
"#,
    ["3", "10"]
);

vb_case!(
    primitive_alias_dictionary_lookup,
    r#"
Imports NameValue = System.Collections.Generic.Dictionary(Of String, Integer)

Module M
    Sub Main()
        Dim map As New NameValue()
        map.Add("alpha", 8)
        map.Add("beta", 3)
        Console.WriteLine(map.Count)
        Console.WriteLine(map("alpha") + map("beta"))
    End Sub
End Module
"#,
    ["2", "11"]
);

vb_case!(
    primitive_alias_math_static_member,
    r#"
Imports MathAlias = System.Math

Module M
    Sub Main()
        Console.WriteLine(MathAlias.Max(4, 9))
        Console.WriteLine(CInt(MathAlias.Round(4.6)))
    End Sub
End Module
"#,
    ["9", "5"]
);

vb_case!(
    primitive_alias_guid_to_string_length,
    r#"
Imports GuidAlias = System.Guid

Module M
    Sub Main()
        Dim value As GuidAlias = GuidAlias.Parse("3F2504E0-4F89-11D3-9A0C-0305E82C3301")
        Console.WriteLine(value.ToString().Length)
        Console.WriteLine(value.GetType().Name)
    End Sub
End Module
"#,
    ["36", "Guid"]
);

vb_case!(
    primitive_alias_global_prefix_respected,
    r#"
Imports IntAlias = Global.System.Int32

Module M
    Sub Main()
        Dim value As IntAlias = 123
        Console.WriteLine(value.GetType().Name)
        Console.WriteLine(CStr(IntAlias.MaxValue))
    End Sub
End Module
"#,
    ["Int32", "2147483647"]
);

vb_case!(
    primitive_alias_parameter_type_contract,
    r#"
Imports Counter = System.Int32

Module M
    Sub Grow(ByVal x As Counter)
        Console.WriteLine(x + 1)
    End Sub

    Sub Main()
        Dim value As Counter = 41
        Grow(value)
    End Sub
End Module
"#,
    ["42"]
);

vb_case!(
    primitive_alias_exception_alias_match,
    r#"
Imports ArgError = System.ArgumentOutOfRangeException

Module M
    Sub Main()
        Try
            Throw New ArgError("index", "bad")
        Catch ex As ArgError
            Console.WriteLine("caught")
            Console.WriteLine(ex.ParamName)
        End Try
    End Sub
    End Module
"#,
    ["caught", "index"]
);

vb_case!(
    primitive_alias_int64_range_and_cast,
    r#"
Imports BigIntAlias = System.Int64

Module M
    Sub Main()
        Dim value As BigIntAlias = BigIntAlias.MaxValue
        Console.WriteLine(value > 0)
        Console.WriteLine(CStr(value \ 3))
    End Sub
End Module
"#,
    ["True", "3074457345618258602"]
);

vb_case!(
    primitive_alias_decimal_precision_with_scale,
    r#"
Imports Money = System.Decimal

Module M
    Sub Main()
        Dim amount As Money = CDec("12.50")
        Dim tax As Money = Money.Round(amount * CDec("0.1"), 2)
        Console.WriteLine(amount.ToString("F2"))
        Console.WriteLine(tax.ToString("F2"))
    End Sub
End Module
"#,
    ["12.50", "1.25"]
);

vb_case!(
    primitive_alias_object_reference_semantics,
    r#"
Imports ObjAlias = System.Object

Module M
    Class Holder
        Public Value As String = "A"
    End Class

    Sub Main()
        Dim left As ObjAlias = New Holder()
        Dim right As ObjAlias = left
        Dim holder As Holder = CType(right, Holder)
        holder.Value = "B"
        Console.WriteLine(CType(left, Holder).Value)
    End Sub
End Module
"#,
    ["B"]
);

vb_case!(
    primitive_alias_nullable_roundtrip_value_semantics,
    r#"
Imports OptionalInt = System.Nullable(Of Integer)

Module M
    Sub Main()
        Dim current As OptionalInt = 7
        Dim empty As OptionalInt = Nothing
        Dim total As Integer = If(current, 0) + If(empty.GetValueOrDefault(0), 0)
        Console.WriteLine(total)
    End Sub
End Module
"#,
    ["7"]
);

vb_case!(
    primitive_alias_timespan_duration_parts,
    r#"
Imports SpanAlias = System.TimeSpan

Module M
    Sub Main()
        Dim elapsed As SpanAlias = SpanAlias.FromHours(1.5)
        Console.WriteLine(elapsed.Hours)
        Console.WriteLine(elapsed.Minutes)
    End Sub
End Module
"#,
    ["1", "30"]
);

vb_case!(
    primitive_alias_datetimeoffset_to_local_parts,
    r#"
Imports OffsetAlias = System.DateTimeOffset

Module M
    Sub Main()
        Dim value As OffsetAlias = OffsetAlias.Parse("2026-12-31T23:45:00+00:00")
        Console.WriteLine(value.Year)
        Console.WriteLine(value.Minute)
    End Sub
End Module
"#,
    ["2026", "45"]
);
