// vybe-test: kotlin/imports/test_import_java_util_arrays
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import java.util.Arrays
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = intArrayOf(3, 1, 2)
            Arrays.sort(a)
            __check((a.joinToString(",")).toString(), "1,2,3")
        }
