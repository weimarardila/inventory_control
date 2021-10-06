# Inventory Control

## EOQ multi-item with Storage Limitations

The model manages n (>1) items, where the items are competing for limited storage capacity.


## Solution method

Step 1. Compute the unconstrained values of the order quantities based on the EOQ equation.

Step 2. Check if the unconstrained values optimal values satisfy the storage constraint. Otherwise, go to the next step.

Step 3. Use the Lagrange multipliers method to determine the constrained optimal values of the order quantities.


The initial value of Lagrange multipliers is usually set equal to zero, and the decrement is set to reasonable value. Those values can be adjusted to enable a better level of accuracy in the calculations.

The VBA app for MS Excel have four commands buttons: Table, EOQ, Random, and Clear.

Table: ask the user for the required input to create a template.

EOQ: calculate the optimal order quantity for each item.

Random: generate a random example.

Clear: clear output. 





