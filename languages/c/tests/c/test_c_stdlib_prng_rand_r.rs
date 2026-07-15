use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn rand_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { srand(42); int r1 = rand(); srand(42); int r2 = rand(); printf(\"%d\", r1 == r2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn rand_r_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 199506L\n#include <stdlib.h>\nint main() { unsigned int seed1 = 42; unsigned int seed2 = 42; int r1 = rand_r(&seed1); int r2 = rand_r(&seed2); printf(\"%d %d\", r1 == r2, seed1 != 42); return 0; }"
        ),
        vec!["1 1"]
    );
}
#[test]
fn random_srandom_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <stdlib.h>\nint main() { srandom(42); long r1 = random(); srandom(42); long r2 = random(); printf(\"%d\", r1 == r2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn initstate_setstate() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <stdlib.h>\nint main() { char state[256]; initstate(42, state, 256); long r1 = random(); setstate(state); long r2 = random(); printf(\"%d\", r1 != r2); return 0; }"
        ),
        vec!["1"]
    );
} // r2 is the next in sequence since state is mutated
#[test]
fn drand48_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { srand48(42); double r1 = drand48(); srand48(42); double r2 = drand48(); printf(\"%d\", r1 == r2 && r1 >= 0.0 && r1 < 1.0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn erand48_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { unsigned short xsubi[3] = {1,2,3}; double r = erand48(xsubi); printf(\"%d\", r >= 0.0 && r < 1.0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn lrand48_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { srand48(42); long r1 = lrand48(); srand48(42); long r2 = lrand48(); printf(\"%d\", r1 == r2 && r1 >= 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn nrand48_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { unsigned short xsubi[3] = {1,2,3}; long r = nrand48(xsubi); printf(\"%d\", r >= 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mrand48_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { srand48(42); long r1 = mrand48(); srand48(42); long r2 = mrand48(); printf(\"%d\", r1 == r2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn jrand48_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { unsigned short xsubi[3] = {1,2,3}; long r = jrand48(xsubi); printf(\"%d\", r == r); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn seed48_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { unsigned short seed[3] = {1,2,3}; unsigned short *old = seed48(seed); printf(\"%d\", old != NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn lcong48_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { unsigned short param[7] = {1,2,3,4,5,6,7}; lcong48(param); double r = drand48(); printf(\"%d\", r >= 0.0 && r < 1.0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn rand_max_check() {
    assert_eq!(
        run_c("#include <stdlib.h>\nint main() { printf(\"%d\", RAND_MAX >= 32767); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn rand_no_seed() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int r = rand(); /* default seed is 1 */ srand(1); int r2 = rand(); printf(\"%d\", r == r2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn rand_r_independence() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 199506L\n#include <stdlib.h>\nint main() { unsigned int s1 = 1, s2 = 2; int r1 = rand_r(&s1); int r2 = rand_r(&s2); printf(\"%d\", r1 != r2); return 0; }"
        ),
        vec!["1"]
    );
} // extremely likely
#[test]
fn srandom_random_range() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <stdlib.h>\nint main() { srandom(100); long r = random(); printf(\"%d\", r >= 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn random_large_state() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <stdlib.h>\nint main() { char state[8]; initstate(1, state, 8); long r = random(); printf(\"%d\", r >= 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn rand_multiple_calls() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { srand(10); int r1 = rand(); int r2 = rand(); printf(\"%d\", r1 != r2); return 0; }"
        ),
        vec!["1"]
    );
} // highly probable
#[test]
fn drand48_multiple_calls() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { srand48(10); double r1 = drand48(); double r2 = drand48(); printf(\"%d\", r1 != r2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn srand48_resets_state() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE\n#include <stdlib.h>\nint main() { srand48(1); drand48(); srand48(1); double r2 = drand48(); srand48(1); double r3 = drand48(); printf(\"%d\", r2 == r3); return 0; }"
        ),
        vec!["1"]
    );
}
