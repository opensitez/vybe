! vybe-test: fortran/arrays/allocatable_assignment_from_elemental_array_call
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer, parameter :: n = 4
    integer, parameter :: indices(*) = [(i, i = 1, n)]
    real, allocatable :: w(:)
    allocate(w(n))
    w = sample_window(indices, n)
    if (abs((w(1)) - 0.25) > 1.0e-6) then
    print *, "FAIL: want [0.25] got [", w(1), "]"
    stop 1
end if
    if (abs((w(2)) - 0.5) > 1.0e-6) then
    print *, "FAIL: want [0.5] got [", w(2), "]"
    stop 1
end if
    if (abs((w(3)) - 0.75) > 1.0e-6) then
    print *, "FAIL: want [0.75] got [", w(3), "]"
    stop 1
end if
    if ((w(4)) /= 1) then
    print *, "FAIL: want [1] got [", w(4), "]"
    stop 1
end if
contains
    elemental pure function sample_window(i, n) result(w)
        integer, intent(in) :: i, n
        real :: w
        w = real(i) / real(n)
    end function sample_window
end program test
