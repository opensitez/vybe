// vybe-test: kotlin/imports/test_import_unused_but_parsed
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.collections.HashSet
        import kotlin.collections.HashMap
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = HashSet<Int>()
            val b = HashMap<String, Int>()
            a.add(1)
            b["x"] = 2
            __check((a.size).toString(), "1")
            __check((b["x"]).toString(), "2")
        }
