! vybe-test: fortran/coarrays/co_broadcast_character
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    character(len=20) :: msg
    if (this_image() == 1) msg = 'hello from 1'
    call co_broadcast(msg, source_image=1)
    print *, trim(msg)
end program test
