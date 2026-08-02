// vybe-test: kotlin/imports/test_import_with_class_scope_not_allowed
// origin: languages/kotlin/tests/kotlin/test_imports.rs

class Host {
            import kotlin.math.abs
            fun norm(v: Int): Int = abs(v)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Host().norm(-7)).toString(), "7")
        }
