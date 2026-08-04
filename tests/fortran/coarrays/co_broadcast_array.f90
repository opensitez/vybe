! vybe-test: fortran/coarrays/co_broadcast_array
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: a(5)
    if (this_image() == 1) a = [1, 2, 3, 4, 5]
    call co_broadcast(a, source_image=1)
    print *, a(3)
end program test
