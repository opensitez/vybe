use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: Complete programs
// ═══════════════════════════════════════════════════════════

#[test]
fn hello_world() {
    let out = run_prints("program hello\nprint *, \"Hello, World!\"\nend program hello\n");
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn sum_1_to_n() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 100\ns = s + i\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["5050"]);
}

#[test]
fn factorial_5() {
    let out = run_prints(
        "program t\ninteger :: i, f\nf = 1\ndo i = 1, 5\nf = f * i\nend do\nprint *, f\nend program t\n",
    );
    assert_eq!(out, vec!["120"]);
}

#[test]
fn fibonacci_10() {
    let out = run_prints(
        "program t\ninteger :: i, a, b, tmp\na = 0\nb = 1\ndo i = 1, 10\ntmp = a + b\na = b\nb = tmp\nend do\nprint *, a\nend program t\n",
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn celsius_to_fahrenheit() {
    let out = run_prints(
        "program t\nreal :: c, f\nc = 100.0\nf = c * 9.0 / 5.0 + 32.0\nprint *, f\nend program t\n",
    );
    assert_eq!(out, vec!["212"]);
}

#[test]
fn power_of_two_table() {
    let out = run_prints(
        "program t\ninteger :: i, p\np = 1\ndo i = 0, 3\nprint *, p\np = p * 2\nend do\nend program t\n",
    );
    assert_eq!(out, vec!["1", "2", "4", "8"]);
}

#[test]
fn min_of_three() {
    let out = run_prints("program t\nprint *, min(min(5, 3), 7)\nend program t\n");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn max_of_three() {
    let out = run_prints("program t\nprint *, max(max(5, 3), 7)\nend program t\n");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn average_computation() {
    let out = run_prints(
        "program t\nreal :: avg\navg = (10 + 20 + 30) / 3.0\nprint *, avg\nend program t\n",
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn swap_variables() {
    let out = run_prints(
        "program t\ninteger :: a, b, tmp\na = 10\nb = 20\ntmp = a\na = b\nb = tmp\nprint *, a\nprint *, b\nend program t\n",
    );
    assert_eq!(out, vec!["20", "10"]);
}

#[test]
fn count_down() {
    let out = run_prints(
        "program t\ninteger :: i\ni = 5\ndo while (i > 0)\nprint *, i\ni = i - 1\nend do\nend program t\n",
    );
    assert_eq!(out, vec!["5", "4", "3", "2", "1"]);
}

#[test]
fn triangle_area() {
    let out = run_prints(
        "program t\nreal :: base, height, area\nbase = 10.0\nheight = 5.0\narea = 0.5 * base * height\nprint *, area\nend program t\n",
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn circle_circumference() {
    let out = run_prints(
        "program t\nreal, parameter :: PI = 3.14159\nreal :: r, c\nr = 5.0\nc = 2.0 * PI * r\nprint *, c\nend program t\n",
    );
    // 2 * 3.14159 * 5 = 31.4159
    assert!(out[0].starts_with("31.4"));
}

#[test]
fn fizzbuzz() {
    let out = run_prints(
        "program t\ninteger :: i\ndo i = 1, 15\nif (mod(i, 15) == 0) then\nprint *, \"FizzBuzz\"\nelse if (mod(i, 3) == 0) then\nprint *, \"Fizz\"\nelse if (mod(i, 5) == 0) then\nprint *, \"Buzz\"\nelse\nprint *, i\nend if\nend do\nend program t\n",
    );
    assert_eq!(
        out,
        vec![
            "1", "2", "Fizz", "4", "Buzz", "Fizz", "7", "8", "Fizz", "10", "11", "Fizz", "13",
            "14", "FizzBuzz",
        ],
    );
}

#[test]
fn sum_even_numbers() {
    let out = run_prints(
        "program t\ninteger :: i, s\ns = 0\ndo i = 1, 10\nif (mod(i, 2) == 0) then\ns = s + i\nend if\nend do\nprint *, s\nend program t\n",
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn quadratic_discriminant() {
    let out = run_prints(
        "program t\nreal :: a, b, c, d\na = 1.0\nb = -5.0\nc = 6.0\nd = b**2 - 4.0*a*c\nif (d >= 0.0) then\nprint *, \"real roots\"\nelse\nprint *, \"complex roots\"\nend if\nend program t\n",
    );
    assert_eq!(out[0].to_lowercase(), "real roots");
}

#[test]
fn full_program_internal_procedures() {
    let out = run_prints(
        "program full_program_internal_procedures\n    integer :: x\n    x = 7\n    print *, adjust(x)\n    print *, total(x)\ncontains\n    integer function adjust(v)\n        integer, intent(in) :: v\n        adjust = v + 3\n    end function adjust\n\n    integer function total(v)\n        integer, intent(in) :: v\n        integer :: i\n        total = 0\n        do i = 1, v\n            total = total + i\n        end do\n    end function total\nend program full_program_internal_procedures\n",
    );
    assert_eq!(out, vec!["10", "28"]);
}

#[test]
fn full_program_module_state_and_calling() {
    let out = run_prints(
        "module full_program_counter\n    integer :: steps = 0\ncontains\n    subroutine advance(by)\n        integer, intent(in) :: by\n        steps = steps + by\n    end subroutine advance\n\n    subroutine reset_counter()\n        steps = 0\n    end subroutine reset_counter\nend module full_program_counter\n\nprogram full_program_module_state_and_calling\n    use full_program_counter\n    call advance(4)\n    call advance(3)\n    print *, steps\n    call reset_counter()\n    print *, steps\nend program full_program_module_state_and_calling\n",
    );
    assert_eq!(out, vec!["7", "0"]);
}

#[test]
fn gcd_iterative() {
    let out = run_prints(
        "program t\ninteger :: a, b, tmp\na = 48\nb = 18\ndo while (b /= 0)\ntmp = b\nb = mod(a, b)\na = tmp\nend do\nprint *, a\nend program t\n",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn is_even_or_odd() {
    let out = run_prints(
        "program t\ninteger :: n = 7\nif (mod(n, 2) == 0) then\nprint *, \"even\"\nelse\nprint *, \"odd\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["odd"]);
}

#[test]
fn absolute_difference() {
    let out = run_prints("program t\nprint *, abs(5 - 12)\nend program t\n");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn string_greeting() {
    let out = run_prints(
        "program t\ncharacter(len=30) :: name\nname = \"Fortran\"\nprint *, \"Hello, \", name\nend program t\n",
    );
    assert!(out[0].contains("Fortran"));
}

#[test]
fn fft_unit_impulse_baseline() {
    let out = run_prints(
        "module fft_probe_module\n    implicit none\n    integer, parameter :: dp = kind(1.0d0)\n    real(dp), parameter :: PI = 4.0_dp * atan(1.0_dp)\ncontains\n    pure subroutine bit_reverse(x)\n        complex(dp), intent(inout) :: x(:)\n        integer :: n, i, j, k\n        complex(dp) :: tmp\n\n        n = size(x)\n        j = 0\n        do i = 1, n - 1\n            k = n / 2\n            do while (j >= k)\n                j = j - k\n                k = k / 2\n            end do\n            j = j + k\n            if (i < j) then\n                tmp = x(i + 1)\n                x(i + 1) = x(j + 1)\n                x(j + 1) = tmp\n            end if\n        end do\n    end subroutine bit_reverse\n\n    subroutine fft(x)\n        complex(dp), intent(inout) :: x(:)\n        integer :: n, stride, half, i, j\n        real(dp) :: angle\n        complex(dp) :: w, wn, tmp\n\n        n = size(x)\n        call bit_reverse(x)\n        stride = 1\n        do while (stride < n)\n            half = stride\n            stride = stride * 2\n            angle = -2.0_dp * PI / stride\n            wn = cmplx(cos(angle), sin(angle), dp)\n            do i = 1, n, stride\n                w = cmplx(1.0_dp, 0.0_dp, dp)\n                do j = 0, half - 1\n                    tmp = w * x(i + j + half)\n                    x(i + j + half) = x(i + j) - tmp\n                    x(i + j) = x(i + j) + tmp\n                    w = w * wn\n                end do\n            end do\n        end do\n    end subroutine fft\nend module fft_probe_module\n\nprogram t\n    use fft_probe_module\n    implicit none\n    complex(dp) :: x(4)\n    integer :: i\n\n    x(1) = cmplx(1.0_dp, 0.0_dp, dp)\n    x(2) = cmplx(0.0_dp, 0.0_dp, dp)\n    x(3) = cmplx(0.0_dp, 0.0_dp, dp)\n    x(4) = cmplx(0.0_dp, 0.0_dp, dp)\n\n    call fft(x)\n    do i = 1, 4\n        print *, nint(real(x(i))), nint(aimag(x(i)))\n    end do\nend program t\n",
    );
    assert_eq!(out, ["1 0", "1 0", "1 0", "1 0"]);
}

#[test]
fn fft_unit_impulse_call_path_runtime() {
    let out = run_prints(
        "module fft_probe_module\n    implicit none\n    integer, parameter :: dp = kind(1.0d0)\n    real(dp), parameter :: PI = 4.0_dp * atan(1.0_dp)\ncontains\n    pure subroutine bit_reverse(x)\n        complex(dp), intent(inout) :: x(:)\n        integer :: n, i, j, k\n        complex(dp) :: tmp\n\n        n = size(x)\n        j = 0\n        do i = 1, n - 1\n            k = n / 2\n            do while (j >= k)\n                j = j - k\n                k = k / 2\n            end do\n            j = j + k\n            if (i < j) then\n                tmp = x(i + 1)\n                x(i + 1) = x(j + 1)\n                x(j + 1) = tmp\n            end if\n        end do\n    end subroutine bit_reverse\n\n    subroutine fft(x)\n        complex(dp), intent(inout) :: x(:)\n        integer :: n, stride, half, i, j\n        real(dp) :: angle\n        complex(dp) :: w, wn, tmp\n\n        n = size(x)\n        call bit_reverse(x)\n        stride = 1\n        do while (stride < n)\n            half = stride\n            stride = stride * 2\n            angle = -2.0_dp * PI / stride\n            wn = cmplx(cos(angle), sin(angle), dp)\n            do i = 1, n, stride\n                w = cmplx(1.0_dp, 0.0_dp, dp)\n                do j = 0, half - 1\n                    tmp = w * x(i + j + half)\n                    x(i + j + half) = x(i + j) - tmp\n                    x(i + j) = x(i + j) + tmp\n                    w = w * wn\n                end do\n            end do\n        end do\n    end subroutine fft\nend module fft_probe_module\n\nprogram t\n    use fft_probe_module\n    implicit none\n    complex(dp) :: x(4)\n\n    x(1) = cmplx(1.0_dp, 0.0_dp, dp)\n    x(2) = cmplx(0.0_dp, 0.0_dp, dp)\n    x(3) = cmplx(0.0_dp, 0.0_dp, dp)\n    x(4) = cmplx(0.0_dp, 0.0_dp, dp)\n\n    call fft(x)\n    print *, 1\nend program t\n",
    );
    assert_eq!(out, ["1"]);
}

#[test]
fn fft_unit_impulse_real_only_runtime() {
    let out = run_prints(
        "module fft_probe_module\n    implicit none\n    integer, parameter :: dp = kind(1.0d0)\n    real(dp), parameter :: PI = 4.0_dp * atan(1.0_dp)\ncontains\n    pure subroutine bit_reverse(x)\n        complex(dp), intent(inout) :: x(:)\n        integer :: n, i, j, k\n        complex(dp) :: tmp\n\n        n = size(x)\n        j = 0\n        do i = 1, n - 1\n            k = n / 2\n            do while (j >= k)\n                j = j - k\n                k = k / 2\n            end do\n            j = j + k\n            if (i < j) then\n                tmp = x(i + 1)\n                x(i + 1) = x(j + 1)\n                x(j + 1) = tmp\n            end if\n        end do\n    end subroutine bit_reverse\n\n    subroutine fft(x)\n        complex(dp), intent(inout) :: x(:)\n        integer :: n, stride, half, i, j\n        real(dp) :: angle\n        complex(dp) :: w, wn, tmp\n\n        n = size(x)\n        call bit_reverse(x)\n        stride = 1\n        do while (stride < n)\n            half = stride\n            stride = stride * 2\n            angle = -2.0_dp * PI / stride\n            wn = cmplx(cos(angle), sin(angle), dp)\n            do i = 1, n, stride\n                w = cmplx(1.0_dp, 0.0_dp, dp)\n                do j = 0, half - 1\n                    tmp = w * x(i + j + half)\n                    x(i + j + half) = x(i + j) - tmp\n                    x(i + j) = x(i + j) + tmp\n                    w = w * wn\n                end do\n            end do\n        end do\n    end subroutine fft\nend module fft_probe_module\n\nprogram t\n    use fft_probe_module\n    implicit none\n    complex(dp) :: x(4)\n    integer :: i\n\n    x(1) = cmplx(1.0_dp, 0.0_dp, dp)\n    x(2) = cmplx(0.0_dp, 0.0_dp, dp)\n    x(3) = cmplx(0.0_dp, 0.0_dp, dp)\n    x(4) = cmplx(0.0_dp, 0.0_dp, dp)\n\n    call fft(x)\n    do i = 1, 4\n        print *, nint(real(x(i)))\n    end do\nend program t\n",
    );
    assert_eq!(out, ["1", "1", "1", "1"]);
}

#[test]
fn fft_unit_impulse_imag_only_runtime() {
    let out = run_prints(
        "module fft_probe_module\n    implicit none\n    integer, parameter :: dp = kind(1.0d0)\n    real(dp), parameter :: PI = 4.0_dp * atan(1.0_dp)\ncontains\n    pure subroutine bit_reverse(x)\n        complex(dp), intent(inout) :: x(:)\n        integer :: n, i, j, k\n        complex(dp) :: tmp\n\n        n = size(x)\n        j = 0\n        do i = 1, n - 1\n            k = n / 2\n            do while (j >= k)\n                j = j - k\n                k = k / 2\n            end do\n            j = j + k\n            if (i < j) then\n                tmp = x(i + 1)\n                x(i + 1) = x(j + 1)\n                x(j + 1) = tmp\n            end if\n        end do\n    end subroutine bit_reverse\n\n    subroutine fft(x)\n        complex(dp), intent(inout) :: x(:)\n        integer :: n, stride, half, i, j\n        real(dp) :: angle\n        complex(dp) :: w, wn, tmp\n\n        n = size(x)\n        call bit_reverse(x)\n        stride = 1\n        do while (stride < n)\n            half = stride\n            stride = stride * 2\n            angle = -2.0_dp * PI / stride\n            wn = cmplx(cos(angle), sin(angle), dp)\n            do i = 1, n, stride\n                w = cmplx(1.0_dp, 0.0_dp, dp)\n                do j = 0, half - 1\n                    tmp = w * x(i + j + half)\n                    x(i + j + half) = x(i + j) - tmp\n                    x(i + j) = x(i + j) + tmp\n                    w = w * wn\n                end do\n            end do\n        end do\n    end subroutine fft\nend module fft_probe_module\n\nprogram t\n    use fft_probe_module\n    implicit none\n    complex(dp) :: x(4)\n    integer :: i\n\n    x(1) = cmplx(1.0_dp, 0.0_dp, dp)\n    x(2) = cmplx(0.0_dp, 0.0_dp, dp)\n    x(3) = cmplx(0.0_dp, 0.0_dp, dp)\n    x(4) = cmplx(0.0_dp, 0.0_dp, dp)\n\n    call fft(x)\n    do i = 1, 4\n        print *, nint(aimag(x(i)))\n    end do\nend program t\n",
    );
    assert_eq!(out, ["0", "0", "0", "0"]);
}
