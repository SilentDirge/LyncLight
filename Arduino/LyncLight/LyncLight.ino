#include "Color.h"

/*
RGB LED Info:
LEDfd
red = 2.6v
green = 3.6v
blue = 3.3v

230ma total (5 leds, 3 channels @ 46ma each):
red = 10ohm
green = 6ohm
blue = 7ohm
*/

Color currentColor = { 0.0f, 0.0f, 0.0f };
Color nextColor = { 0.0f, 0.0f, 0.0f };

bool lowBrite = false;

void transitionTo(Color col) {
  nextColor = col;
}

void applyColor(Color col) {
  float scalar;
  if (lowBrite) {
    scalar = 127;
  } else {
    scalar = 255;
  }
  analogWrite(redPin, scalar - col.r * scalar);
  analogWrite(greenPin, scalar - col.g * scalar);
  analogWrite(bluePin, scalar - col.b * scalar);
}

float transitionRate = 0.001f;
bool lightShowSpeedFast = false;

void updateColorTransitions() {
  float rate = transitionRate;
  if (lightShowSpeedFast) {
    rate *= 10;
  }
  currentColor.r = currentColor.r + (nextColor.r - currentColor.r) * rate;
  currentColor.g = currentColor.g + (nextColor.g - currentColor.g) * rate;
  currentColor.b = currentColor.b + (nextColor.b - currentColor.b) * rate;
  applyColor(currentColor);
} 

void setup() {
  // initialize digital pin 13 as an output.
  pinMode(redPin, OUTPUT);
  pinMode(greenPin, OUTPUT);
  pinMode(bluePin, OUTPUT);
  
  digitalWrite(redPin, HIGH);
  digitalWrite(greenPin, HIGH);
  digitalWrite(bluePin, HIGH);
  
  currentColor = black;
  nextColor = white;
  applyColor(white);
  
  Serial.begin(9600);
}

void blink(int pin) {
  digitalWrite(pin, LOW);
  delay(1000);              // wait for a second
  digitalWrite(pin, HIGH);
  delay(1000);              // wait for a second
}

bool lightShow = false;
int currentLightShowColor = 0;
void applyLightShowColor() {
  switch (currentLightShowColor) {
    case 0:
      transitionTo(red);
     break;
    case 1:
      transitionTo(green);
     break;     
    case 2:
      transitionTo(blue);
     break;     
    case 3:
      transitionTo(yellow);
     break;     
    case 4:
      transitionTo(magenta);
     break;     
    case 5:
      transitionTo(orange);
     break;          
    case 6:
      transitionTo(purple);
     break;               
  }
}

bool lightShowChangeColor() {
 return fabs(currentColor.r - nextColor.r) < 0.01 &&
        fabs(currentColor.g - nextColor.g) < 0.01 &&
        fabs(currentColor.b - nextColor.b) < 0.01;
}

void loop() {
//  blink(3);
//  blink(5);
//  blink(9);
/*  for (int i = 0; i < 255; ++i) {
      analogWrite(3, i);
      delay(10);
  }*/
  
/*  applyColor(red);
  delay(200);
  applyColor(green);
  delay(200);
  applyColor(blue);
  delay(200);*/
  
  if (lightShow) {
    if (lightShowChangeColor()) {
      currentLightShowColor = (currentLightShowColor + 1) % 7;
      applyLightShowColor();      
    }
  }

  updateColorTransitions();
  delay(10);
}

void serialEvent() {
  while (Serial.available()) {
    char inChar = (char)Serial.read();
    
    lightShow = false;
    if (inChar == 'r') {
      transitionTo(red);
    } else if (inChar == 'g') {
      transitionTo(green);
    } else if (inChar == 'b') {
      transitionTo(blue);
    } else if (inChar == 'y') {
      transitionTo(yellow);
    } else if (inChar == 'o') {
      transitionTo(orange);
    } else if (inChar == 'm') {
      transitionTo(magenta);
    } else if (inChar == 'p') {
      transitionTo(purple);
    } else if (inChar == 'w') {
      transitionTo(white);      
    } else if (inChar == 'x') {
      lowBrite = !lowBrite;
    } else if (inChar == 'z') {
      lightShow = true;
      currentLightShowColor = 0;
      applyLightShowColor();
    } else if (inChar == 'f') {
      lightShowSpeedFast = !lightShowSpeedFast;
      lightShow = true;
      currentLightShowColor = 0;
      applyLightShowColor();      
    }
  }
}
