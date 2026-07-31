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
    module_alias_string_builder_append_chain,
    r#"
Imports Txt = System.Text

Module M
    Sub Main()
        Dim sb As New Txt.StringBuilder()
        sb.Append("A").Append("-").Append("B")
        Console.WriteLine(sb.ToString())
    End Sub
End Module
"#,
    ["A-B"]
);

vb_case!(
    module_alias_string_builder_length_after_append,
    r#"
Imports Txt = System.Text

Module M
    Sub Main()
        Dim sb As New Txt.StringBuilder("seed")
        sb.Append("123")
        Console.WriteLine(sb.Length)
    End Sub
End Module
"#,
    ["7"]
);

vb_case!(
    module_alias_text_encoding_ascii,
    r#"
Imports EncodingAlias = System.Text.Encoding

Module M
    Sub Main()
        Dim bytes() As Byte = EncodingAlias.UTF8.GetBytes("go")
        Console.WriteLine(bytes.Length)
        Console.WriteLine(CStr(bytes(0)))
    End Sub
End Module
"#,
    ["2", "103"]
);

vb_case!(
    module_alias_path_get_extension,
    r#"
Imports Paths = System.IO.Path

Module M
    Sub Main()
        Console.WriteLine(Paths.GetExtension("document.backup.txt"))
    End Sub
End Module
"#,
    [".txt"]
);

vb_case!(
    module_alias_path_file_name_no_ext,
    r#"
Imports Paths = System.IO.Path

Module M
    Sub Main()
        Console.WriteLine(Paths.GetFileNameWithoutExtension("data.tar.gz"))
    End Sub
End Module
"#,
    ["data.tar"]
);

vb_case!(
    module_alias_io_directory_name,
    r#"
Imports IOAlias = System.IO

Module M
    Sub Main()
        Dim part As String = IOAlias.Path.GetFileNameWithoutExtension("dir/archive.json")
        Console.WriteLine(part)
    End Sub
End Module
"#,
    ["archive"]
);

vb_case!(
    module_alias_generic_list_growth_and_index,
    r#"
Imports Gen = System.Collections.Generic

Module M
    Sub Main()
        Dim values As New Gen.List(Of Integer)()
        values.Add(2)
        values.Add(4)
        values.Add(6)
        Console.WriteLine(values.Count)
        Console.WriteLine(values(1))
    End Sub
End Module
"#,
    ["3", "4"]
);

vb_case!(
    module_alias_generic_dictionary_contains_key,
    r#"
Imports Ctx = System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Ctx.Dictionary(Of String, Integer)()
        map("one") = 1
        map("two") = 2
        Console.WriteLine(map.ContainsKey("two"))
        Console.WriteLine(map("two"))
    End Sub
End Module
"#,
    ["True", "2"]
);

vb_case!(
    module_alias_math_maximum,
    r#"
Imports MathAlias = System.Math

Module M
    Sub Main()
        Console.WriteLine(MathAlias.Max(3, 17))
        Console.WriteLine(CInt(MathAlias.Min(4, 9)))
    End Sub
End Module
"#,
    ["17", "4"]
);

vb_case!(
    module_alias_system_convert,
    r#"
Imports Conv = System.Convert

Module M
    Sub Main()
        Dim text As String = "255"
        Console.WriteLine(Conv.ToInt32(text))
        Console.WriteLine(Conv.ToBoolean("True"))
    End Sub
End Module
"#,
    ["255", "True"]
);

vb_case!(
    module_alias_garbage_collection_stats,
    r#"
Imports GCs = System.GC

Module M
    Sub Main()
        GCs.Collect()
        Console.WriteLine("ok")
        Console.WriteLine(CStr(GCs.MaxGeneration))
    End Sub
End Module
"#,
    ["ok", "2"]
);

vb_case!(
    module_alias_array_clone_preserved,
    r#"
Imports Arr = System

Module M
    Sub Main()
        Dim source() As Integer = {1, 2, 3}
        Dim copy() As Integer = Arr.ConvertAll(source, Function(x) x + 2)
        copy(1) = 9
        Console.WriteLine(source(1))
        Console.WriteLine(copy(1))
    End Sub
End Module
"#,
    ["2", "9"]
);

vb_case!(
    module_alias_datetime_parts,
    r#"
Imports Dates = System

Module M
    Sub Main()
        Dim now As Dates.DateTime = Dates.DateTime.Parse("2026-02-03T00:00:00")
        Console.WriteLine(now.Month)
        Console.WriteLine(CStr(now.DayOfWeek))
    End Sub
End Module
"#,
    ["2", "Tuesday"]
);

vb_case!(
    module_alias_linq_contains,
    r#"
Imports LinqAlias = System.Linq.Enumerable

Module M
    Sub Main()
        Dim values As Integer() = {1, 2, 3, 4}
        Console.WriteLine(LinqAlias.Any(values, Function(v) v = 3))
        Console.WriteLine(LinqAlias.All(values, Function(v) v > 0))
    End Sub
End Module
"#,
    ["True", "True"]
);

vb_case!(
    module_alias_text_upper_case_behavior,
    r#"
Imports Txt = System.Text
Imports CultureText = System.Globalization

Module M
    Sub Main()
        Dim comparer = CultureText.CultureInfo.InvariantCulture
        Dim input As New Txt.StringBuilder("abc")
        Console.WriteLine(input.ToString().ToUpper(comparer))
    End Sub
End Module
"#,
    ["ABC"]
);

vb_case!(
    module_alias_string_isnullorempty,
    r#"
Imports Strings = System.String

Module M
    Sub Main()
        Console.WriteLine(Strings.IsNullOrEmpty(""))
        Console.WriteLine(Strings.IsNullOrWhiteSpace("   "))
    End Sub
End Module
"#,
    ["True", "True"]
);

vb_case!(
    module_alias_regex_match,
    r#"
Imports RegexAlias = System.Text.RegularExpressions.Regex

Module M
    Sub Main()
        Console.WriteLine(RegexAlias.IsMatch("abc-123", "^[a-z]+-\d+$"))
        Console.WriteLine(RegexAlias.Replace("a-b-c", "-", "_"))
    End Sub
End Module
"#,
    ["True", "a_b_c"]
);

vb_case!(
    module_alias_queue_roundtrip,
    r#"
Imports QAlias = System.Collections.Generic

Module M
    Sub Main()
        Dim q As New QAlias.Queue(Of Integer)()
        q.Enqueue(4)
        q.Enqueue(8)
        Console.WriteLine(q.Count)
        Console.WriteLine(q.Dequeue())
        Console.WriteLine(q.Peek())
        Console.WriteLine(q.Count)
    End Sub
End Module
"#,
    ["2", "4", "8", "1"]
);

vb_case!(
    module_alias_convert_typecasts,
    r#"
Imports Conv = System.Convert

Module M
    Sub Main()
        Dim n As Integer = Conv.ToInt32("12")
        Dim b As Boolean = Conv.ToBoolean("true")
        Dim d As Double = Conv.ToDouble("4")
        Console.WriteLine(n)
        Console.WriteLine(b)
        Console.WriteLine(CInt(d))
    End Sub
End Module
"#,
    ["12", "True", "4"]
);

vb_case!(
    module_alias_text_upper_with_culture,
    r#"
Imports Texts = System.Text
Imports Culture = System.Globalization.CultureInfo

Module M
    Sub Main()
        Dim title As New Texts.StringBuilder("ß")
        Dim culture = Culture.InvariantCulture
        Console.WriteLine(title.ToString().ToUpper(culture))
    End Sub
End Module
"#,
    ["SS"]
);

vb_case!(
    module_alias_stopwatch_running_state,
    r#"
Imports Sw = System.Diagnostics.Stopwatch

Module M
    Sub Main()
        Dim watch As Sw = Sw.StartNew()
        watch.Stop()
        Console.WriteLine(CStr(watch.IsRunning))
        Console.WriteLine(CStr(watch.ElapsedMilliseconds >= 0))
    End Sub
End Module
"#,
    ["False", "True"]
);

vb_case!(
    module_alias_environment_version,
    r#"
Imports Env = System.Environment

Module M
    Sub Main()
        Console.WriteLine(CStr(Env.Version.Major >= 0))
        Console.WriteLine(CStr(Env.Version.MajorRevision > -1))
    End Sub
End Module
"#,
    ["True", "True"]
);

vb_case!(
    module_alias_uri_parts,
    r#"
Imports UriAlias = System.Uri

Module M
    Sub Main()
        Dim link As New UriAlias("https://example.com/search?q=vb")
        Console.WriteLine(link.Scheme)
        Console.WriteLine(link.Host)
        Console.WriteLine(link.Query)
    End Sub
End Module
"#,
    ["https", "example.com", "?q=vb"]
);

vb_case!(
    module_alias_bitconverter_length,
    r#"
Imports BC = System.BitConverter

Module M
    Sub Main()
        Dim bytes() As Byte = BC.GetBytes(CShort(12))
        Console.WriteLine(bytes.Length)
        Console.WriteLine(CStr(BC.ToInt16(bytes, 0)))
    End Sub
End Module
"#,
    ["2", "12"]
);

vb_case!(
    module_alias_version_compare_to,
    r#"
Imports V = System.Version

Module M
    Sub Main()
        Dim one As New V(1, 2, 3, 4)
        Dim two As New V(1, 2, 4, 4)
        Console.WriteLine(CStr(one.CompareTo(two)))
        Console.WriteLine(one.ToString())
    End Sub
End Module
"#,
    ["-1", "1.2.3.4"]
);
