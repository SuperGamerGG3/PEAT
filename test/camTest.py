import cv2

# Open webcam
cam = cv2.VideoCapture(0)

# Read first frame
ret, frame1 = cam.read()
ret, frame2 = cam.read()

while True:

    # Difference between frames
    diff = cv2.absdiff(frame1, frame2)

    # Convert to grayscale
    gray = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)

    # Blur to reduce noise
    blur = cv2.GaussianBlur(gray, (5, 5), 0)

    # Bright areas = motion
    _, thresh = cv2.threshold(blur, 20, 255, cv2.THRESH_BINARY)

    # Find motion areas
    contours, _ = cv2.findContours(
        thresh,
        cv2.RETR_TREE,
        cv2.CHAIN_APPROX_SIMPLE
    )

    motion_detected = False

    for contour in contours:

        # Ignore tiny movements
        if cv2.contourArea(contour) < 1000:
            continue

        motion_detected = True

        # Draw rectangle around motion
        x, y, w, h = cv2.boundingRect(contour)

        cv2.rectangle(
            frame1,
            (x, y),
            (x + w, y + h),
            (0, 255, 0),
            2
        )

    if motion_detected:
        print("MOTION DETECTED")

    # Show webcam
    cv2.imshow("PEAT Surveillance", frame1)

    # Shift frames forward
    frame1 = frame2
    ret, frame2 = cam.read()

    # Quit with Q
    if cv2.waitKey(1) & 0xFF == ord("q"):
        break

cam.release()
cv2.destroyAllWindows()