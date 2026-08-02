// vybe-test: kotlin/named_arguments/test_named_arguments_object_factory_style
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

class Item(val a: Int, val b: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val i = Item(a = 3, b = 4)
            __check((i.a + i.b).toString(), "7")
        }
