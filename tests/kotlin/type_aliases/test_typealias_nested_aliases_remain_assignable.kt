// vybe-test: kotlin/type_aliases/test_typealias_nested_aliases_remain_assignable
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias BaseLabel = String
        typealias UserLabel = BaseLabel
        typealias DisplayLabel = UserLabel

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source: BaseLabel = "admin"
            val alias: DisplayLabel = source
            val roundTrip: UserLabel = alias
            __check((alias).toString(), "admin")
            __check((roundTrip).toString(), "admin")
        }
