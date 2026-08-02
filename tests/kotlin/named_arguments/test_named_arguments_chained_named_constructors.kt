// vybe-test: kotlin/named_arguments/test_named_arguments_chained_named_constructors
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

class Holder(val a: Int, val b: Int)
        class Container(val left: Holder, val title: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Container(left = Holder(a = 1, b = 2), title = "t")
            __check((c.left.a + c.left.b).toString(), "3")
            __check((c.title).toString(), "t")
        }
