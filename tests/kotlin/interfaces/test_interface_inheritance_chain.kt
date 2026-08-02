// vybe-test: kotlin/interfaces/test_interface_inheritance_chain
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Root { fun name(): String }
interface Mid : Root { fun suffix(): String = ".mid" }
interface Leaf : Mid { fun tail(): String = ".leaf" }
class Thing : Leaf { override fun name(): String = "t" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val t = Thing()
__check((t.name() + t.suffix() + t.tail()).toString(), "t.mid.leaf") }
