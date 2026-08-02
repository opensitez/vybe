// vybe-test: kotlin/function_overloads/test_overload_for_class_and_top_level_names
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun echo(v: Int): String = "top" + v
        class Host {
            fun echo(v: Int): String = "member" + v
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Host()
            __check((echo(1)).toString(), "top1")
            __check((h.echo(2)).toString(), "member2")
        }
