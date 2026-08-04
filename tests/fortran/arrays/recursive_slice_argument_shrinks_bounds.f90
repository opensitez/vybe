! vybe-test: fortran/arrays/recursive_slice_argument_shrinks_bounds
! origin: languages/fortran/tests/fortran/test_arrays.rs
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 3, 2, 1 ]

recursive subroutine trim_tail(arr)
    integer, intent(in) :: arr(:)
    integer :: n
    n = size(arr)
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((n) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", n, "]"
        stop 1
    end if
    if (n <= 1) return
    call trim_tail(arr(2:))
end subroutine trim_tail

program test
    integer :: a(3) = [1, 2, 3]
    call trim_tail(a)
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program test
