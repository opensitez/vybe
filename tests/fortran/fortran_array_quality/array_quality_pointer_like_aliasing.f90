! vybe-test: fortran/fortran_array_quality/array_quality_pointer_like_aliasing
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_pointer_like_aliasing
    integer, target, dimension(3) :: source
    integer, pointer :: head
    source = (/ 11, 22, 33 /)
    head => source(2)
    if ((head) /= 22) then
    print *, "FAIL: want [22] got [", head, "]"
    stop 1
end if
end program array_quality_pointer_like_aliasing
