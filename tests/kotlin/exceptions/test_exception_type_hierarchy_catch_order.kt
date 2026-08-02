// vybe-test: kotlin/exceptions/test_exception_type_hierarchy_catch_order
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

class BaseError : Exception("base")
        class DerivedError : BaseError()

        fun main() {
            try {
                throw DerivedError()
            } catch (e: DerivedError) {
                println("derived")
            } catch (e: BaseError) {
                println("base")
            } catch (e: Exception) {
                println("general")
            }
        }

