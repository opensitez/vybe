' vybe-test: vb/vb_xml_descendant_axis/xml_descendant_axis
' origin: languages/vb/tests/vb/test_vb_xml_descendant_axis.rs

Module M
    Sub Main()
        Dim xml = <Root>
                      <Container>
                          <Item>One</Item>
                      </Container>
                      <Item>Two</Item>
                  </Root>
                  
        ' XML descendant axis operator ... returns all matching descendants at any level
        Dim items = xml...<Item>
        
        For Each item In items
            Console.WriteLine(item.Value)
        Next
    End Sub
End Module
