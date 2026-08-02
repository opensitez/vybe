// vybe-test: kotlin/class_delegation/test_delegate_with_custom_extension_function_usage
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Text { fun text(): String }

        class Source : Text {
            override fun text() = "value"
        }

        class Delegate(delegate: Text) : Text by delegate

        fun Text.enhancedSuffix(): String = text() + "!"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val d = Delegate(Source())
            __check((d.enhancedSuffix()).toString(), "value!")
        }
