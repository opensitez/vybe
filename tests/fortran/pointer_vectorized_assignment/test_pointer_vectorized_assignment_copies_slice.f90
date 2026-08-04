! vybe-test: fortran/pointer_vectorized_assignment/test_pointer_vectorized_assignment_copies_slice
! origin: languages/fortran/tests/fortran/test_pointer_vectorized_assignment.rs

program test_pointer_vectorized_assignment
    integer, target :: src(3)
    integer, target :: dst(3)
    integer, pointer :: psrc(:)
    integer, pointer :: pdst(:)

    src = (/1, 2, 3/)
    psrc => src
    pdst => dst
    pdst = psrc
    if ((pdst(2)) /= 2) then
    print *, "FAIL: want [2] got [", pdst(2), "]"
    stop 1
end if
end program test_pointer_vectorized_assignment
