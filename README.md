# Random Name Generator

The Random Name Generator (RNG) is an Excel spreadsheet that draws random names from a weighted pool to display to the user. RNG also uses VBA macros to enable extra functionality to the spreadsheet.

## Drawing Names

There are two ways of drawing names:

1. Drawing a single name
2. Drawing multiple names

By drawing multiple names, the user has the ability to paste the list of names onto another sheet so it can be easily printed onto paper.

Once names have been drawn, the name will be on cooldown and will not be drawn for x amount of draws. It is rare to have a name drawn back to back if there are a lot of names in the pool.

## Adding Names

The user has the ability to add names to the pool along with a weight, which gives them a chance to draw that name at some point in the future.

## Weight Options

The user has many options to change the weight of the drawn name, below are the following options:

- Set weight to 1
- Set weight to half of its original value
- Subtract weight by 1
- Set weight to a percentage of its original weight (0.001% to 100%)
- Do nothing

## Resetting the Cooldown

The user has the ability to change the cooldown on the draws, the higher the cooldown means the name will not be drawn for longer. The user also has the ability to force a reset, which will allow all names to be drawn again.

## Future Additions

There are many planned additions to this spreadsheet, below are some of the following:

- **Force draw names** - If there are names that have *force draw* enabled, there will be another pool of names in which those names will be drawn first before the other names will be drawn.
- **Probability of a name getting drawn** - The user will enter another form which will ask them to enter a name. Once entered, the program will display the probability of the name getting drawn
