// vybe-test: kotlin/interfaces/test_interface_default_and_override
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Source { fun text(): String = "base" }
class OverrideSource : Source { override fun text(): String = "child" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val s: Source = OverrideSource()
__check((s.text()).toString(), "child") }
