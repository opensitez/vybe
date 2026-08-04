! vybe-test: fortran/select_case_extended/case_on_product
! origin: languages/fortran/tests/fortran/test_select_case_extended.rs
program t
integer :: vybe_check_i = 0
character(len=13) :: vybe_check_w(1) = [ "large product" ]
integer :: a = 3, b = 4
select case (a * b)
case (1:10)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("small product") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "small product", "]"
    stop 1
end if
case (11:20)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("large product") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "large product", "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
