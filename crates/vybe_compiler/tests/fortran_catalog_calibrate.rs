#[path = "fortran/helpers.rs"]
mod helpers;

use helpers::run_prints;

macro_rules! cal {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            match std::panic::catch_unwind(|| run_prints($src)) {
                Ok(out) => eprintln!("OK|{}|{:?}", stringify!($name), out),
                Err(_) => eprintln!("FAIL|{}", stringify!($name)),
            }
        }
    };
}

cal!(
    all_true,
    "program t\nlogical :: m(4) = [.true., .true., .true., .true.]\nprint *, all(m)\nend program t\n"
);
cal!(
    allocated,
    "program t\ninteger, allocatable :: a(:)\nprint *, allocated(a)\nallocate(a(2))\nprint *, allocated(a)\nend program t\n"
);
cal!(
    bge,
    "program t\nprint *, bge('abc', 'abc')\nend program t\n"
);
cal!(
    bessel_j0,
    "program t\nprint *, nint(bessel_j0(0.0)*100)\nend program t\n"
);
cal!(
    count_dim1,
    "program t\ninteger :: a(2,3) = reshape([1,2,3,4,5,6],[2,3])\nprint *, count(a > 3, dim=1)\nend program t\n"
);
cal!(dble, "program t\nprint *, nint(dble(7))\nend program t\n");
cal!(digits, "program t\nprint *, digits(1.0)\nend program t\n");
cal!(dim, "program t\nprint *, dim(10, 3)\nend program t\n");
cal!(
    dot_product,
    "program t\ninteger :: a(3) = [1,2,3]\ninteger :: b(3) = [4,5,6]\nprint *, dot_product(a,b)\nend program t\n"
);
cal!(
    dprod,
    "program t\nprint *, nint(dprod(2.0d0, 3.0d0))\nend program t\n"
);
cal!(
    dshiftl,
    "program t\nprint *, dshiftl(14, 2, 4)\nend program t\n"
);
cal!(
    eoshift,
    "program t\ninteger :: a(4) = [1,2,3,4]\nprint *, eoshift(a, 1)\nend program t\n"
);
cal!(
    epsilon,
    "program t\nprint *, exponent(epsilon(1.0))\nend program t\n"
);
cal!(
    erf,
    "program t\nprint *, nint(erf(0.0)*100)\nend program t\n"
);
cal!(
    exp,
    "program t\nprint *, nint(exp(0.0)*100)\nend program t\n"
);
cal!(
    exponent,
    "program t\nprint *, exponent(16.0)\nend program t\n"
);
cal!(
    fraction,
    "program t\nprint *, nint(fraction(3.75)*100)\nend program t\n"
);
cal!(
    gamma,
    "program t\nprint *, nint(gamma(2.0)*100)\nend program t\n"
);
cal!(huge, "program t\nprint *, huge(0) > 0\nend program t\n");
cal!(iall, "program t\nprint *, iall(7)\nend program t\n");
cal!(
    ibits,
    "program t\nprint *, ibits(170, 1, 4)\nend program t\n"
);
cal!(leadz, "program t\nprint *, leadz(1)\nend program t\n");
cal!(
    lgamma,
    "program t\nprint *, nint(lgamma(2.0)*100)\nend program t\n"
);
cal!(logical, "program t\nprint *, logical(1)\nend program t\n");
cal!(
    log_fn,
    "program t\nprint *, nint(log(1.0)*100)\nend program t\n"
);
cal!(
    command_argument_count,
    "program t\nprint *, command_argument_count()\nend program t\n"
);
cal!(
    cpu_time,
    "program t\nreal :: t\ncall cpu_time(t)\nprint *, nint(t*100)\nend program t\n"
);
cal!(
    date_and_time,
    "program t\ninteger :: dt(8)\ncall date_and_time(values=dt)\nprint *, dt(1)\nend program t\n"
);
cal!(
    get_command,
    "program t\ncharacter(len=32) :: cmd\ninteger :: stat\nstat = get_command(cmd)\nprint *, stat\nend program t\n"
);
cal!(
    image_index,
    "program t\nprint *, image_index(1)\nend program t\n"
);
cal!(
    is_iostat_end,
    "program t\nprint *, is_iostat_end(-1)\nend program t\n"
);
cal!(
    lbound_whole,
    "program t\ninteger :: a(4)\nprint *, lbound(a)\nend program t\n"
);
cal!(
    cshift_dim,
    "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\nprint *, cshift(m, 1, dim=2)\nend program t\n"
);
cal!(
    findloc_dim1,
    "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\nprint *, findloc(m, 5, dim=1)\nend program t\n"
);
cal!(
    acos,
    "program t\nprint *, nint(acos(1.0)*100)\nend program t\n"
);
cal!(
    achar,
    "program t\nprint *, iachar(achar(72))\nend program t\n"
);
cal!(
    adjustl,
    "program t\ncharacter(len=10) :: s = '   data'\nprint *, len_trim(adjustl(s))\nend program t\n"
);
cal!(
    aimag,
    "program t\ncomplex :: z = (4.0, -3.0)\nprint *, nint(aimag(z))\nend program t\n"
);
cal!(aint, "program t\nprint *, aint(3.9)\nend program t\n");
cal!(anint, "program t\nprint *, anint(3.5)\nend program t\n");
cal!(
    any_dim1,
    "program t\nlogical :: m(2,2) = reshape([.false.,.true.,.false.,.false.],[2,2])\nlogical :: c(2)\nc = any(m, dim=1)\nprint *, c(1)\nend program t\n"
);
cal!(
    asin,
    "program t\nprint *, nint(asin(0.0)*100)\nend program t\n"
);
cal!(
    atan2,
    "program t\nprint *, nint(atan2(1.0,1.0)*100)\nend program t\n"
);
