! vybe-test: fortran/coarrays/co_broadcast_single_image
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: x = 99
    call co_broadcast(x, source_image=1)
    print *, x
end program test
