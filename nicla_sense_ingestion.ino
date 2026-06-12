#include "Arduino_BHY2.h"
#include "ArduinoBLE.h"

#define FREQUENCY_HZ     50        // 50 Hz — enough to capture writing motion
#define INTERVAL_MS      (1000 / FREQUENCY_HZ)
#define CONVERT_G_TO_MS2 9.80665f

// BLE — UUIDs must match the Python config exactly
BLEService imuService("19b10000-e8f2-537e-4f6c-d104768a1214");
BLECharacteristic accelChar(
  "19b10001-e8f2-537e-4f6c-d104768a1214",
  BLERead | BLENotify,
  6  // 3 x int16 = 6 bytes
);

SensorXYZ accel(SENSOR_ID_ACC);
static unsigned long last_interval_ms = 0;

void setup() {
  Serial.begin(115200);

  BHY2.begin(NICLA_I2C);
  accel.begin();

  if (!BLE.begin()) {
    Serial.println("BLE failed to start!");
    while (1);
  }

  BLE.setLocalName("NiclaSenseME");
  BLE.setAdvertisedService(imuService);
  imuService.addCharacteristic(accelChar);
  BLE.addService(imuService);
  BLE.advertise();

  Serial.println("BLE advertising started");
}

void loop() {
  BHY2.update();
  BLE.poll();

  if (millis() - last_interval_ms >= INTERVAL_MS) {
    last_interval_ms = millis();

    // Convert raw values to m/s² (same scaling your original sketch used)
    float ax = (accel.x() * 8.0 / 32768.0) * CONVERT_G_TO_MS2;
    float ay = (accel.y() * 8.0 / 32768.0) * CONVERT_G_TO_MS2;
    float az = (accel.z() * 8.0 / 32768.0) * CONVERT_G_TO_MS2;

    // Print to Serial Monitor so you can still see raw values
    Serial.print(ax); Serial.print(", ");
    Serial.print(ay); Serial.print(", ");
    Serial.println(az);

    // Pack as int16 (scale by 100 to preserve 2 decimal places)
    int16_t bx = (int16_t)(ax * 100);
    int16_t by = (int16_t)(ay * 100);
    int16_t bz = (int16_t)(az * 100);

    uint8_t packet[6];
    memcpy(packet + 0, &bx, 2);
    memcpy(packet + 2, &by, 2);
    memcpy(packet + 4, &bz, 2);

    // Only send if a central is connected
    BLEDevice central = BLE.central();
    if (central && central.connected()) {
      accelChar.writeValue(packet, 6);
    }
  }
}