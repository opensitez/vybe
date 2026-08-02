// vybe-test: kotlin/named_arguments/test_named_arguments_all_defaults_can_be_overridden_by_name
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun build(prefix: String = "a", middle: String = "b", suffix: String = "c"): String {
            return prefix + middle + suffix
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((build()).toString(), "abc")
            __check((build(suffix = "Z")).toString(), "abZ")
            __check((build(middle = "Y", prefix = "X")).toString(), "XYc")
        }
