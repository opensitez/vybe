// vybe-test: kotlin/interfaces/test_interface_property_override
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Readable { val value: String }
class A : Readable { override val value = "alpha" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val r: Readable = A()
__check((r.value).toString(), "alpha") }
