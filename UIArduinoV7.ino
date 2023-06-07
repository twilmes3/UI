int mode = 1; // Initial mode
float raw = 0.0;
float Vin = 5.0;
float Vout = 0.0;
float R11 = 100.0;
float R22 = 0.0;
float buffer = 0.0;
float calibration = 0.0;
float vout = 0.0;
float vin = 0.0;
float R1 = 10000.0; // resistance of R1 (100K) -see text!
float R2 = 2200.0; // resistance of R2 (2.2K) - see text!
int value = 0;
float sumresistance = 0.0;


void setup() {
  Serial.begin(115200);
  pinMode(0, INPUT);
  pinMode(9, OUTPUT);
//  tone(9, 262,100);
//  delay(150);
//  tone(9, 349,100);
//  delay(150);
//  tone(9, 392,100);
//  delay(40);
  digitalWrite(9, LOW);
  pinMode(13, OUTPUT);
}

void loop() {
   if (Serial.available()) {
     char receivedChar = Serial.read();
     if (receivedChar == '1') {
       mode = 1;
     } 
     else if (receivedChar == '2') {
       mode = 2;
     }
     else if (receivedChar == '3') {
       mode = 3;
     }
     else if (receivedChar == '4') {
       mode = 4;
     }
     else if (receivedChar == '5') {
       mode = 5;
     }
   }

  if (mode == 1) {

      value = analogRead(A0);
      vout = (value * 5.0) / 1024.0; // see text
      vin = vout / (R2 / (R1 + R2));
      Serial.println(vin);
      delay(100);
  } else if (mode == 2) {
      value = analogRead(A2);
      vout = (value * 5.0) / 1024.0; // see text
      vout = 5.0 - vout;
      Serial.println(vout);
      delay(100);
  }
    else if  (mode == 3){
      raw = analogRead(A2);
      buffer = raw * Vin;
      Vout = (buffer) / 1024.0;
      buffer = (Vin / Vout) - 1.0;
      R22 = (R11 * buffer) - calibration;
      Serial.println(R22);
      delay(100);
    }
    else if (mode == 4){
      raw = analogRead(A2);
      buffer = raw * Vin;
      Vout = (buffer) / 1024.0;
      buffer = (Vin / Vout) - 1.0;
      R22 = (R11 * buffer) - calibration;
      Serial.println(R22);
      delay(100);
      if(R22 < 10)
      {
        tone(9, 349,350);
      }
      else
      {
        digitalWrite(9, LOW);
      }
    }
    else if (mode == 5){
      calibration = 0.0;
      for (int i = 0; i < 10; i++){
      raw = analogRead(A2);
      buffer = raw * Vin;
      Vout = (buffer) / 1024.0;
      buffer = (Vin / Vout) - 1.0;
      R22 = (R11 * buffer);
      sumresistance = sumresistance + R22;
      if (i == 9){
        calibration = sumresistance/10.0;
        if (calibration >10.0){
                  calibration = 0.0;
                  int fail = 0;
                  Serial.println(fail);
                  delay(100);
        }
        else if (calibration <= 10.0){
        int pass = 1;
        Serial.println(pass);
        delay(100);
        }
        delay(1000);
        mode = 1;
    }
    }
}
}
