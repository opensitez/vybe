// vybe-test: csharp/csharp_enum_guid_version/guid_parse_round_trips_stable_input_string
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var text = "00112233-4455-6677-8899-aabbccddeeff"; __Check((System.Guid.Parse(text).ToString()).ToString(), "00112233-4455-6677-8899-aabbccddeeff");
