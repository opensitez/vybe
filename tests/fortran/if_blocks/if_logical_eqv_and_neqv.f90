! vybe-test: fortran/if_blocks/if_logical_eqv_and_neqv
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(1) = [ "neqv" ]
logical :: a = .true., b = .false.
if (a .eqv. b) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("eqv") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "eqv", "]"
    stop 1
end if
end if
if ((a .neqv. b)) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("neqv") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "neqv", "]"
    stop 1
end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
