# Matrix Rain Lab Assignment

## Objective

Build a Matrix-style digital rain animation with a hidden message that gets revealed as rain falls over it.

**Difficulty:** Intermediate
**Time Estimate:** 2-3 hours
**Prerequisites:** Basic JavaScript, HTML Canvas basics, Arrays

---

## Final Result

When complete, you should have:

- Green characters falling down the screen in columns
- A large hidden smiley ":)" in the center
- Characters that turn yellow when they overlap the smiley
- A fading trail effect behind falling characters
- Smooth, continuous animation at ~30 FPS

---

## Setup (5 minutes)

### Task 1: Create the HTML Structure

1. Create a new HTML file called `matrix-rain.html`
2. Add a `<canvas>` element with id "matrix"
3. Add basic CSS to remove margins and set background to black
4. Set canvas to display as block and hide overflow

### Task 2: Initialize Canvas in JavaScript

1. Get a reference to the canvas element
2. Get the 2D drawing context
3. Set canvas width to 1920 pixels
4. Set canvas height to 1080 pixels

**Test:** Open in browser - you should see a black screen.

---

## Part 1: Character Setup (10 minutes)

### Task 3: Define Your Character Set

1. Create a string variable containing Matrix-style characters
2. Include these katakana characters: `ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ`
3. Add numbers 0-9 to the character set
4. Create a variable for font size (start with 20)

**Hint:** Store these as constants since they won't change.

### Task 4: Calculate Grid Dimensions

1. Calculate how many columns fit across the canvas width
2. Divide canvas width by font size
3. Store this value in a variable

**Think:** If width is 1920 and fontSize is 20, how many columns do you get?

---

## Part 2: The Drops System (15 minutes)

### What is a "Drop"?

A **drop** is just a number representing the Y position (vertical position) of the rain in one column.

**Example:** If `drops[5] = 10`, it means column 5's rain is at Y position 10 (in character heights, not pixels).

Think of it as "where the rain has dropped to" in each column.

### Task 5: Create the Drops Array

1. Create an empty array to store drop positions (we'll call this array `drops`)
2. Loop from 0 to the number of columns
3. For each column, initialize a random starting Y position
4. Random position should be between 0 and (canvas height / font size)
5. Store this position in the drops array

**Result:** An array where `drops[i]` = the Y position for column i

**Think:** Why random starting positions instead of all at zero?

### Task 6: Understand the Drops Array

Before moving on, make sure you understand:

- What does each element in the array represent?
- What do the values mean (position in characters, not pixels)?
- How will this array change over time?

---

## Part 3: Draw the Hidden Message (20 minutes)

### Task 7: Create a Function to Draw the Message

1. Create a function called `drawHiddenMessage`
2. Inside it, set the font size to 20 times larger than the rain characters
3. Set font family to monospace
4. Set fill color to a very faint green with very low alpha (almost invisible)
5. Set text alignment to center
6. Draw the text ":)" at the center of the canvas

**Canvas Functions You'll Need:**

- `ctx.font` - Sets font size and family (format: "20px monospace")
- `ctx.fillStyle` - Sets the fill color (use 'rgba(red, green, blue, alpha)')
- `ctx.textAlign` - Sets text alignment ('left', 'center', or 'right')
- `ctx.textBaseline` - Sets vertical alignment ('top', 'middle', or 'bottom')
- `ctx.fillText(text, x, y)` - Draws text at position (x, y)

**Test:** Call this function once and check your browser. Can you barely see the smiley?

### Task 8: Create a Detection Function

1. Create a function called `isInsideMessage` that takes x and y coordinates
2. Use `getImageData` to read the pixel at position (x, y)
3. Check if the pixel has any green value
4. Return true if green > 0, false otherwise

**Hint:** Pixel data is an array with format [R, G, B, A]. Green is at index 1.

**Think:** Why are we checking for green? What color is our hidden message?

---

## Part 4: The Animation Loop (30-45 minutes)

### Task 9: Create the Draw Function Structure

1. Create a function called `draw`
2. This function will be called repeatedly to create animation
3. For now, just create the empty function

### Task 10: Add the Trail Effect

Inside your `draw` function:

1. Set fill style to semi-transparent black (RGBA with low alpha)
2. Draw a rectangle covering the entire canvas
3. Alpha value should be around 0.05

**Think:** Why semi-transparent instead of fully opaque? What visual effect does this create?

**Key Concept:** This creates the "trail" effect. Each frame:
- Old characters get slightly dimmed by this black overlay
- New characters are drawn at new positions (one per column)
- Old characters from previous frames gradually fade away
- You never redraw old characters - they just fade naturally!

### Task 11: Redraw the Hidden Message

1. Call your `drawHiddenMessage` function
2. This needs to happen every frame to keep the message persistent

### Task 12: Set Font for Rain Characters

1. Set the font to your fontSize variable with monospace family
2. This font will be used for all the falling characters

### Task 13: Loop Through Each Column

1. Create a for loop that iterates through all drops
2. Remember: Each drop represents ONE column of rain
3. For each iteration, you'll draw ONE character in that column at its current position

**Important:** You're drawing one character per column per frame, not redrawing the entire trail!

### Task 14: Pick a Random Character

Inside your loop (once per column):

1. Generate a random index between 0 and the length of your character set
2. Use this index to pick a random character from your string
3. This character will be drawn at THIS column's current position

**Hint:** Use `Math.random()`, multiply by character set length, and use `Math.floor()`.

**Remember:** Each column gets one new random character per frame!

### Task 15: Calculate Position

Still inside the loop:

1. Calculate the X position by multiplying column index by font size
2. Calculate the Y position by multiplying the drop's value by font size

**Think:** Why multiply by fontSize? What units are the drop values in?

### Task 16: Determine Color Based on Message Overlap

1. Call your `isInsideMessage` function with the calculated x and y positions
2. If it returns true, set fill style to yellow (#ffff00)
3. If it returns false, set fill style to green (#0f0)

**This is the key step!** This is what reveals the hidden message.

### Task 17: Draw the Character

1. Use `fillText` to draw your random character at position (x, y)

### Task 18: Move the Drop Down

1. Increment the drop's position by 1
2. This makes it fall down one character height for the next frame

### Task 19: Reset Drops When Off-Screen

1. Check if the y position (in pixels) is greater than canvas height
2. Also add a random condition (Math.random() > 0.975)
3. If both conditions are true, reset the drop to position 0

**Think:** Why the random condition? What would happen without it?

---

## Part 5: Start the Animation (5 minutes)

### Task 20: Create the Animation Loop

1. Use `setInterval` to call your `draw` function repeatedly
2. Set interval to 33 milliseconds (approximately 30 FPS)

**Test:** Open in browser. You should see:

- Green characters falling down
- A yellow smiley revealed in the center
- Trailing effect behind characters

---

## Part 6: Testing & Debugging (15 minutes)

### Checkpoint 1: Basic Animation

- [ ] Characters are falling downward
- [ ] Characters are random (not all the same)
- [ ] Animation is smooth
- [ ] No errors in console

### Checkpoint 2: Hidden Message

- [ ] Smiley is barely visible in the center
- [ ] Characters turn yellow over the smiley
- [ ] The smiley shape is clearly revealed by yellow characters
- [ ] Green characters appear everywhere else

### Checkpoint 3: Visual Effects

- [ ] Characters have a trailing/fading effect
- [ ] Trails are not too long or too short
- [ ] Drops reset smoothly when reaching bottom
- [ ] Drops don't all fall in sync

### Common Issues to Check:

- If no smiley appears: Check the alpha value in drawHiddenMessage
- If all characters are the same: Check random character selection
- If no trails: Check the alpha value in the fade rectangle
- If characters disappear instantly: Make sure you're not using clearRect
- If drops sync up: Check random initialization and reset condition

---

## Part 7: Enhancements (Optional Challenges)

### Challenge 1: Change the Message

Modify your code to display "HELLO" instead of ":)"

### Challenge 2: Multiple Messages

Display two different messages at different positions on screen

### Challenge 3: Different Colors

Make the rain blue and the revealed message white

### Challenge 4: Speed Control

Add a variable to control animation speed (frames per second)

### Challenge 5: Adjustable Trail Length

Make the trail length easily adjustable with one variable

### Challenge 6: Different Character Sets

Create multiple character sets (Latin, numbers only, symbols) and randomly switch between them

### Challenge 7: Brightness Variation

Make characters at the leading edge of each drop brighter than trailing characters

### Challenge 8: Wave Effect

Make the drops move slightly left and right as they fall (sine wave pattern)

---

## Deliverables

When you're done, you should have:

1. A working `matrix-rain.html` file
2. Clean, readable code
3. Comments explaining your logic
4. All checkpoints passing

---

## Reflection Questions

After completing the lab, answer these:

1. **How does the drops array work?**
   - What does each element represent?
   - How is it updated each frame?

2. **How does the trail effect work?**
   - Why use semi-transparent rectangles?
   - What would happen with different alpha values?

3. **How does message detection work?**
   - Why can we use pixel color to detect the message?
   - What are limitations of this approach?

4. **What determines animation smoothness?**
   - What's the relationship between setInterval delay and FPS?
   - How could you make it smoother or choppier?

5. **How would you optimize this for performance?**
   - What's the most expensive operation?
   - What could you do less frequently?

---

## Success Criteria

**Minimum Viable Product (Pass):**

- Green characters falling in columns
- Hidden message visible through color change
- Basic trail effect
- Smooth animation with no crashes

**Good Implementation:**

- All of the above, plus:
- Clean, commented code
- Proper variable naming
- Consistent style
- At least one enhancement challenge completed

**Excellent Implementation:**

- All of the above, plus:
- Multiple enhancement challenges completed
- Creative customizations
- Optimized performance
- Well-organized code structure

---

## Resources

**If you get stuck:**

1. Check the browser console for errors
2. Use `console.log()` to debug values
3. Test each function individually
4. Refer to `MATRIX_RAIN_TUTORIAL.md` for hints (but try without it first!)
5. Check `matrix-rain.html` for the complete solution (last resort)

**Canvas API References:**

- fillRect() - draws rectangles
- fillText() - draws text
- getImageData() - reads pixel colors
- setInterval() - creates animation loops

---

## Time Management

Suggested time allocation:

- Setup: 5 min
- Character setup: 10 min
- Drops system: 15 min
- Hidden message: 20 min
- Animation loop: 45 min
- Testing: 15 min
- Enhancements: 30+ min (optional)

**Total Core Time:** ~2 hours
**With Enhancements:** 2.5-3 hours

---

## Tips for Success

1. **Test frequently** - Don't write everything before testing
2. **One step at a time** - Complete each task before moving on
3. **Use console.log()** - Debug by printing values
4. **Comment as you go** - Explain your logic for future you
5. **Experiment** - Change values and see what happens
6. **Take breaks** - Step away if stuck for 15+ minutes
7. **Read error messages** - They usually tell you exactly what's wrong
8. **Start simple** - Get basic version working before adding features

---

## What You'll Learn

**Technical Skills:**

- Canvas API manipulation
- Animation loop patterns
- Array-based systems
- Pixel sampling techniques
- Transparency effects
- Random number generation
- Performance considerations

**Concepts:**

- Frame-based animation
- State management (drops array)
- Visual effects with math
- Optimization strategies
- Debugging techniques

---

## Next Steps

After completing this lab:

1. Try building a particle system
2. Create a game using similar techniques
3. Experiment with other Canvas effects
4. Study the tutorial MD for deeper understanding
5. Share your customized version!

---

Good luck! Remember: the goal is to learn by doing. Struggle a bit before checking the solution - that's where the real learning happens. 🟢
