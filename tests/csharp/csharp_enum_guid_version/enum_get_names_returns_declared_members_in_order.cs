// vybe-test: csharp/csharp_enum_guid_version/enum_get_names_returns_declared_members_in_order
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

enum State { Idle, Running, Done } foreach (var name in System.Enum.GetNames(typeof(State))) Console.WriteLine(name);
