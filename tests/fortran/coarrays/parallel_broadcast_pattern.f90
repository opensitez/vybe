! vybe-test: fortran/coarrays/parallel_broadcast_pattern
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: seed[*]
    if (this_image() == 1) seed = 42
    call co_broadcast(seed, source_image=1)
    print *, seed
end program test
