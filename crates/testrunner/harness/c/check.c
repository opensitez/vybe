/* Vybe test harness — C.
 *
 * Real C: this compiles with `cc` on its own. Like the COBOL and Fortran
 * harnesses it documents the shape the emitter produces rather than providing
 * a function to splice, because the check has to be inline — see below.
 *
 * WHY THE CHECK IS INLINE AND PER PRINT.
 * Every way of ACCUMULATING output fails under Vybe while working under cc
 * (all measured):
 *
 *   static char g[64] written from a function  ->  undefined is not callable
 *                                                  (__libc_char_to_str)
 *   snprintf(buf + n, ...)  — pointer arithmetic -> js-string.compare: first
 *                                                  arg not a string
 *   strcat(dst, src)                           ->  js-number.toF64: not a number
 *   vsnprintf + va_list                        ->  undefined is not callable
 *
 * So there is no buffer to collect into, and each `printf` is compared where
 * it stands. What DOES work in both: a LOCAL `char t[N]`, `snprintf` writing
 * at offset 0, `strcmp` against a literal, and `printf("%s", t)`.
 *
 * WHY assert(0) AND NOT exit/return.
 * `assert` is the ONLY failure signal that reaches a non-zero status in both:
 *
 *   assert(0)            cc 134   vybex 1    <- the one that works
 *   exit(1) in a block   cc 1     vybex 0
 *   return 3 from main   cc 3     vybex 0
 *   abort()              cc 134   vybex 0
 *
 * `exit_call_is_return = true` (languages/c/src/profile:14) compiles `exit` to
 * a RETURN, and main's return value is not mapped to the process status. Both
 * are real C bugs; until they are fixed `assert` is the only option, and it is
 * enough because the verdict only needs non-zero.
 *
 * The FAIL line is printed BEFORE asserting: an uncaught failure renders as
 * `RuntimeError: [object]` under Vybe, so the values would otherwise be lost.
 */
#include <stdio.h>
#include <string.h>
#include <assert.h>

int main(void) {
    int a[] = {5, 3, 8};

    /* One emitted check, replacing `printf("%d %d %d\n", a[0], a[1], a[2]);`
     * whose expected line is "5 3 8". The braces keep `__t` scoped so repeated
     * checks in one function do not collide. */
    {
        char __t[512];
        snprintf(__t, sizeof(__t), "%d %d %d\n", a[0], a[1], a[2]);
        if (strcmp(__t, "5 3 8\n") != 0) {
            printf("FAIL: want [5 3 8] got [%s]\n", __t);
            assert(0);
        }
    }

    printf("harness ok\n");
    return 0;
}
