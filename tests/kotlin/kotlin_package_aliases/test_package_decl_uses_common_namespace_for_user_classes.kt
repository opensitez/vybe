// vybe-test: kotlin/kotlin_package_aliases/test_package_decl_uses_common_namespace_for_user_classes
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

package demo.core

        class User(val name: String) {
            fun label(): String = name
        }

        fun make(): User = User("Ada")

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((demo.core.make().label()).toString(), "Ada")
        }
