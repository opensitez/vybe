! vybe-test: fortran/coarrays/co_broadcast_integer
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x
    if (this_image() == 1) x = 42
    call co_broadcast(x, source_image=1)
    print *, x
end program test
