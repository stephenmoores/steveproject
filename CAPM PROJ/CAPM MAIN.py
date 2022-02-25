print("Welcome to Steve's CAPM model")

TCKR = input("what is the ticker of the stock you wish to evaluate?   ")
RF = float(input("what is the current risk free rate   "))
BETA = float(input("what is the beta of the stock you wish to evaluate?   "))
RM = float(input("what is the current expected market return?   "))

CAPM1 = float(RM) - float(RF)
CAPM2 = BETA * CAPM1
CAPM = float(RF) + CAPM2
print(f'the expected return of {TCKR} is {round(CAPM*100,4)} percent')

