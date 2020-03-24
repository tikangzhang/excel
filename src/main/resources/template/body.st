 $worksheet:{
 <Worksheet ss:Name="$it.name$">
  <Table ss:ExpandedColumnCount="8" ss:ExpandedRowCount="$it.rowNum$" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="108" ss:DefaultRowHeight="14.25">
   <Row>
    <Cell ss:StyleID="s65"><Data ss:Type="String">$it.header.f0$</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">$it.header.f1$</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">$it.header.f2$</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">$it.header.f3$</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">$it.header.f4$</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">$it.header.f5$</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">$it.header.f6$</Data></Cell>
    <Cell ss:StyleID="s65"><Data ss:Type="String">$it.header.f7$</Data></Cell>
   </Row>
 $it.rows:{
   <Row>
    <Cell><Data ss:Type="String">$it.f0$</Data></Cell>
    <Cell><Data ss:Type="String">$it.f1$</Data></Cell>
    <Cell><Data ss:Type="String">$it.f2$</Data></Cell>
    <Cell><Data ss:Type="String">$it.f3$</Data></Cell>
    <Cell><Data ss:Type="String">$it.f4$</Data></Cell>
    <Cell><Data ss:Type="String">$it.f5$</Data></Cell>
    <Cell><Data ss:Type="String">$it.f6$</Data></Cell>
    <Cell><Data ss:Type="String">$it.f7$</Data></Cell>
   </Row>
 }$
  </Table>
 </Worksheet>
}$