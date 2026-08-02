// vybe-test: kotlin/interfaces/test_interface_static_lookup_like
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Id { fun id(): Int }
class One : Id { override fun id(): Int = 1 }
class Two : Id { override fun id(): Int = 2 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val list: Array<Id> = arrayOf(One(), Two())
__check((list[0].id() + list[1].id()).toString(), "3") }
