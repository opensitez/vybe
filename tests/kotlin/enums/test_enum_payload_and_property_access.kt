// vybe-test: kotlin/enums/test_enum_payload_and_property_access
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Zone(val id: Int) { A(10), B(20), C(30) }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Zone.B.id).toString(), "20") }
