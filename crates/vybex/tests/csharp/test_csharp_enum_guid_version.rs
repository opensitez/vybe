use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(enum_get_names_returns_declared_members_in_order, r#"enum State { Idle, Running, Done } foreach (var name in System.Enum.GetNames(typeof(State))) Console.WriteLine(name);"#, ["Idle", "Running", "Done"]);
csharp_case!(enum_parse_converts_string_to_enum_value, r#"enum State { Idle, Running, Done } Console.WriteLine(System.Enum.Parse(typeof(State), "Running"));"#, ["Running"]);
csharp_case!(enum_try_parse_reports_success_for_valid_name, r#"enum State { Idle, Running } System.Enum.TryParse<State>("Idle", out var value); Console.WriteLine(value);"#, ["Idle"]);
csharp_case!(enum_is_defined_reports_true_for_existing_value, r#"enum State { Idle, Running } Console.WriteLine(System.Enum.IsDefined(typeof(State), "Running"));"#, ["True"]);
csharp_case!(enum_has_flag_detects_enabled_bit_flag, r#"[System.Flags] enum Permission { Read = 1, Write = 2, Execute = 4 } var value = Permission.Read | Permission.Write; Console.WriteLine(value.HasFlag(Permission.Write));"#, ["True"]);
csharp_case!(enum_to_string_formats_combined_flags, r#"[System.Flags] enum Permission { Read = 1, Write = 2 } var value = Permission.Read | Permission.Write; Console.WriteLine(value.ToString());"#, ["Read, Write"]);
csharp_case!(enum_get_underlying_type_reports_int_by_default, r#"enum State { Idle } Console.WriteLine(System.Enum.GetUnderlyingType(typeof(State)).Name);"#, ["Int32"]);
csharp_case!(enum_format_d_outputs_numeric_representation, r#"enum State { Idle = 1 } Console.WriteLine(System.Enum.Format(typeof(State), State.Idle, "D"));"#, ["1"]);
csharp_case!(guid_empty_has_all_zero_text_representation, r#"Console.WriteLine(System.Guid.Empty.ToString());"#, ["00000000-0000-0000-0000-000000000000"]);
csharp_case!(guid_parse_round_trips_stable_input_string, r#"var text = "00112233-4455-6677-8899-aabbccddeeff"; Console.WriteLine(System.Guid.Parse(text).ToString());"#, ["00112233-4455-6677-8899-aabbccddeeff"]);
csharp_case!(guid_constructor_from_string_matches_parse_output, r#"var text = "11111111-2222-3333-4444-555555555555"; Console.WriteLine(new System.Guid(text).ToString());"#, ["11111111-2222-3333-4444-555555555555"]);
csharp_case!(guid_try_parse_reports_true_for_valid_text, r#"var ok = System.Guid.TryParse("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee", out var value); Console.WriteLine(ok); Console.WriteLine(value.ToString());"#, ["True", "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"]);
csharp_case!(guid_try_parse_reports_false_for_invalid_text, r#"var ok = System.Guid.TryParse("bad-guid", out var value); Console.WriteLine(ok);"#, ["False"]);
csharp_case!(version_parse_exposes_major_minor_and_build, r#"var version = System.Version.Parse("2.4.6"); Console.WriteLine(version.Major); Console.WriteLine(version.Minor); Console.WriteLine(version.Build);"#, ["2", "4", "6"]);
csharp_case!(version_to_string_round_trips_original_text, r#"var version = new System.Version(1, 2, 3, 4); Console.WriteLine(version.ToString());"#, ["1.2.3.4"]);
csharp_case!(version_compare_to_reports_negative_for_smaller_version, r#"var left = new System.Version(1, 2); var right = new System.Version(1, 3); Console.WriteLine(left.CompareTo(right));"#, ["-1"]);
csharp_case!(version_equals_reports_true_for_identical_versions, r#"Console.WriteLine(new System.Version(3, 5).Equals(new System.Version(3, 5)));"#, ["True"]);
csharp_case!(enum_get_values_can_be_enumerated, r#"enum State { Idle, Running } foreach (var value in System.Enum.GetValues(typeof(State))) Console.WriteLine(value);"#, ["Idle", "Running"]);
csharp_case!(enum_try_parse_can_ignore_case_when_requested, r#"enum State { Idle, Running } System.Enum.TryParse<State>("running", true, out var value); Console.WriteLine(value);"#, ["Running"]);
csharp_case!(version_revision_defaults_to_negative_one_when_missing, r#"var version = new System.Version(1, 2, 3); Console.WriteLine(version.Revision);"#, ["-1"]);