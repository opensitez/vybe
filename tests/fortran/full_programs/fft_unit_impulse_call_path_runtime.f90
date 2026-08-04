! vybe-test: fortran/full_programs/fft_unit_impulse_call_path_runtime
! origin: languages/fortran/tests/fortran/test_full_programs.rs
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
module fft_probe_module
    implicit none
    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: PI = 4.0_dp * atan(1.0_dp)
contains
    pure subroutine bit_reverse(x)
        complex(dp), intent(inout) :: x(:)
        integer :: n, i, j, k
        complex(dp) :: tmp

        n = size(x)
        j = 0
        do i = 1, n - 1
            k = n / 2
            do while (j >= k)
                j = j - k
                k = k / 2
            end do
            j = j + k
            if (i < j) then
                tmp = x(i + 1)
                x(i + 1) = x(j + 1)
                x(j + 1) = tmp
            end if
        end do
    end subroutine bit_reverse

    subroutine fft(x)
        complex(dp), intent(inout) :: x(:)
        integer :: n, stride, half, i, j
        real(dp) :: angle
        complex(dp) :: w, wn, tmp

        n = size(x)
        call bit_reverse(x)
        stride = 1
        do while (stride < n)
            half = stride
            stride = stride * 2
            angle = -2.0_dp * PI / stride
            wn = cmplx(cos(angle), sin(angle), dp)
            do i = 1, n, stride
                w = cmplx(1.0_dp, 0.0_dp, dp)
                do j = 0, half - 1
                    tmp = w * x(i + j + half)
                    x(i + j + half) = x(i + j) - tmp
                    x(i + j) = x(i + j) + tmp
                    w = w * wn
                end do
            end do
        end do
    end subroutine fft
end module fft_probe_module

program t
    use fft_probe_module
    implicit none
    complex(dp) :: x(4)

    x(1) = cmplx(1.0_dp, 0.0_dp, dp)
    x(2) = cmplx(0.0_dp, 0.0_dp, dp)
    x(3) = cmplx(0.0_dp, 0.0_dp, dp)
    x(4) = cmplx(0.0_dp, 0.0_dp, dp)

    call fft(x)
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((1) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
