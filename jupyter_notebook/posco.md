Input files:
![Input-file](https://res.cloudinary.com/thinhnt/image/upload/v1627440258/uni/inputs_r6kh9w.png)

# POSCO:

- South Korean steel-making company headquartered in Pohang, South Korea

## STEPS IN BALANCE.E21:

- Take items in xnk.xlsx and divide it to three table base on type: E21, B13, A42.
- Merge rows with same ITEM CODE and in column 'Tổng số lượng' using sum as aggregation functions (
  E21: 'Import column',
  B13: 'Re-Export',
  A42: 'Re-Purpose',
  )
- 'Begin' column: pending...
- 'Output for Production', 'Warehouse' and 'Return' columns from sum of 'Weight' column in 'takeout.xlsx', 'iob.xlsx' and 'return.xlsx' ('Warehouse' taken from 'Closing Balance') (If 'Unit' is 'KG' take 'Weight' else take 'Qty')


## Final goal
  1. result by period(Monthly, Quarterly, Half year, year) and type(E21, E31);
  - same as Excel balance sheet;
  - the begin of each period result change as last period balance
  2. Find reason if Ecus balance coulum is minus(-) value;
  - Normally miss imput from DO table, A12 in EXIM, etc  ;
  3. Discrepancy between table;
  - for exemple, import Qty in EXIM but not Take-in;
  - for exemple, Different Qty between table( 100 in take-in but 50 in EXIM);
