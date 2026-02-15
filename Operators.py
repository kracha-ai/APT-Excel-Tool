class Cat:
    furr= "soft"

    def __init__(self, color, weight, name):
        self.color = color
        self.weight = weight
        self.name=name

cat1=Cat("black", 2, "smiley")
cat2=Cat("Gold", 3, "Dumba")

print(f"{cat2.name} weighs {cat2.weight} kg and is in {cat2.color} color")
