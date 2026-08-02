// vybe-test: kotlin/interfaces/test_interface_default_message
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Message { fun msg(): String = "hi" }
class Holder: Message
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val value: Message = Holder()
__check((value.msg()).toString(), "hi") }
