// vybe-test: kotlin/kotlin_string_line_ops/test_string_replace_and_split_ops
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("aa-bb-cc".replaceFirst("-", "/")).toString(), "aa/bb-cc")
            __check(("aa-bb-cc".replace("-", "/")).toString(), "aa/bb/cc")
            val out = "a,b,c".split(",")
            __check((out.size).toString(), "3")
            __check((out[1]).toString(), "b")
            __check(("a,b,c".split(",", limit = 2).size).toString(), "2")
        }
