Set WshShell = CreateObject("WScript.Shell")

' Open Notepad
WshShell.Run "notepad.exe", 1, True

' Wait for Notepad to open
WScript.Sleep 1000

' Type out the recipe
WshShell.SendKeys "A Very Simple Salad That You Can Literally Make At Home"
WshShell.SendKeys "{ENTER}{ENTER}Ingredients:{ENTER}- Lettuce: Choose your preferred type of lettuce, such as romaine, iceberg, or mixed greens, as the base of your salad.{ENTER}- Ranch dressing: Use a small amount of ranch dressing to add creaminess and tangy flavor to the salad. Adjust the quantity according to your taste preferences.{ENTER}- Cheese: A combination of Swiss, mozzarella, and sharp cheddar cheese. This blend of cheeses provides a variety of flavors and textures, with Swiss offering a mild nuttiness, mozzarella adding a smooth and elastic texture, and sharp cheddar providing a bold and tangy taste.{ENTER}- Tomatoes: Include fresh tomatoes for a burst of juiciness and added flavor. You can slice them or dice them based on your preference.{ENTER}{ENTER}Instructions:{ENTER}1. Wash and dry the lettuce leaves. Tear or chop them into bite-sized pieces and place them in a salad bowl.{ENTER}2. Slice or dice the tomatoes and add them to the bowl with the lettuce.{ENTER}3. Sprinkle the cheese blend over the lettuce and tomatoes.{ENTER}4. Drizzle a small amount of ranch dressing over the salad ingredients.{ENTER}5. Use salad tongs or a large spoon to gently toss and mix all the ingredients together, ensuring the dressing and cheese are evenly distributed.{ENTER}6. Once everything is well combined, you can serve the salad immediately.{ENTER}{ENTER}Remember, you can always customize this recipe by adding other vegetables like cucumbers, carrots, or bell peppers, or including protein sources like grilled chicken or chickpeas for a more substantial meal. Adjust the quantities of each ingredient to suit your taste and preferences. Enjoy your delicious homemade salad!"

' Press CTRL+P
WshShell.SendKeys "^p"

' Wait for the print dialog to appear
WScript.Sleep 2000

' Display a message
MsgBox "JUST CHOOSE THE PRINTER DEVICE!!! OR PRESS CANCEL IF YOU ARE SO INCLINED!"
