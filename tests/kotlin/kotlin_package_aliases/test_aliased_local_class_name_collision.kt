// vybe-test: kotlin/kotlin_package_aliases/test_aliased_local_class_name_collision
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.collections.HashSet as Bucket

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = Bucket<Int>()
            left.add(1)
            left.add(2)
            __check((left.size).toString(), "2")
        }
