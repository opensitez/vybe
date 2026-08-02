// vybe-test: kotlin/imports/test_import_package_objects_in_expression
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.sqrt
        import kotlin.math.PI
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((sqrt(PI) * 2).toInt()).toString(), "3")
        }
