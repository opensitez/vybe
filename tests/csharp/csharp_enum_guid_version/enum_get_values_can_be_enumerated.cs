// vybe-test: csharp/csharp_enum_guid_version/enum_get_values_can_be_enumerated
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

enum State { Idle, Running } foreach (var value in System.Enum.GetValues(typeof(State))) Console.WriteLine(value);
