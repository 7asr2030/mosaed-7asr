// المتغيرات العامة
let zip = new JSZip();
let students = [];
let currentStudent = null;
let currentPoseIndex = 0;
let capturedCount = 0;
let faceMesh = null;
let camera = null;

// تسلسل الحركات
const poses = ['center', 'right', 'left', 'up', 'down'];
const poseNames = {
    'center': 'الأمام',
    'right': 'يمين',
    'left': 'يسار',
    'up': 'أعلى',
    'down': 'أسفل'
};

const instructions = {
    'center': 'يرجى النظر للأمام مباشرة',
    'right': 'أدر وجهك قليلاً إلى اليمين',
    'left': 'أدر وجهك قليلاً إلى اليسار',
    'up': 'ارفع وجهك للأعلى قليلاً',
    'down': 'أخفض وجهك للأسفل قليلاً'
};

// معايرة المركز
let centerYaw = 1.0;
let centerPitch = 1.0;
let captureHoldFrames = 0;
const REQUIRED_HOLD_FRAMES = 15; // يجب تثبيت الوجه لـ 15 إطار (نصف ثانية تقريباً)

// عناصر الواجهة
const stepUpload = document.getElementById('step-upload');
const stepStudents = document.getElementById('step-students');
const stepCamera = document.getElementById('step-camera');
const excelInput = document.getElementById('excelFileInput');
const uploadStatus = document.getElementById('uploadStatus');
const studentsList = document.getElementById('studentsList');
const searchInput = document.getElementById('searchInput');
const studentsCount = document.getElementById('studentsCount');
const downloadZipBtn = document.getElementById('downloadZipBtn');

const videoElement = document.getElementById('videoElement');
const canvasElement = document.getElementById('canvasElement');
const canvasCtx = canvasElement.getContext('2d');
const instructionText = document.getElementById('instructionText');
const guideCircle = document.querySelector('.guide-circle');
const captureProgress = document.getElementById('captureProgress');
const currentStudentName = document.getElementById('currentStudentName');
const backToListBtn = document.getElementById('backToListBtn');

// 1. قراءة ملف الإكسل
excelInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    uploadStatus.textContent = "جاري قراءة الملف...";
    uploadStatus.className = "status-msg";

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // قراءة الورقة الأولى
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // تحويل إلى مصفوفة كائنات
            const rawData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            
            parseStudentsData(rawData);
        } catch (error) {
            uploadStatus.textContent = "❌ حدث خطأ في قراءة الملف، تأكد من أنه ملف إكسل صالح.";
            uploadStatus.className = "status-msg error";
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
});

function parseStudentsData(data) {
    // البحث عن أعمدة الاسم والهوية
    let idIndex = -1;
    let nameIndex = -1;
    let headerRow = -1;

    for (let i = 0; i < Math.min(10, data.length); i++) {
        const row = data[i];
        if (!row) continue;
        
        for (let j = 0; j < row.length; j++) {
            const cellVal = String(row[j] || '').trim();
            if (cellVal.includes('رقم الطالب') || cellVal.includes('رقم الهوية') || cellVal === 'الهوية') {
                idIndex = j;
            }
            if (cellVal.includes('اسم الطالب') || cellVal.includes('الاسم') || cellVal === 'الاسم الرباعي') {
                nameIndex = j;
            }
        }
        
        if (idIndex !== -1 && nameIndex !== -1) {
            headerRow = i;
            break;
        }
    }

    if (idIndex === -1 || nameIndex === -1) {
        uploadStatus.textContent = "❌ لم يتم العثور على أعمدة (رقم الهوية) و (اسم الطالب) في الملف.";
        uploadStatus.className = "status-msg error";
        return;
    }

    students = [];
    for (let i = headerRow + 1; i < data.length; i++) {
        const row = data[i];
        if (!row || !row[idIndex]) continue;
        
        let id = String(row[idIndex]).trim();
        // إزالة الفواصل العشرية إن وجدت
        if (id.includes('.')) id = id.split('.')[0];
        // تنظيف من الحروف
        id = id.replace(/\D/g,'');
        
        const name = String(row[nameIndex] || 'غير محدد').trim();
        
        if (id) {
            students.push({ id: id, name: name, done: false });
        }
    }

    if (students.length > 0) {
        uploadStatus.textContent = `✅ تم بنجاح! وجدنا ${students.length} طالب.`;
        uploadStatus.className = "status-msg success";
        studentsCount.textContent = students.length;
        
        setTimeout(() => {
            stepUpload.classList.remove('active');
            stepStudents.classList.add('active');
            renderStudentsList();
            initFaceMesh(); // تهيئة الكاميرا والذكاء الاصطناعي مسبقاً
        }, 1000);
    } else {
        uploadStatus.textContent = "❌ الملف فارغ أو لا يحتوي على بيانات صحيحة.";
        uploadStatus.className = "status-msg error";
    }
}

// 2. عرض وقائمة الطلاب
function renderStudentsList(filter = '') {
    studentsList.innerHTML = '';
    const filtered = students.filter(s => s.name.includes(filter) || s.id.includes(filter));
    
    filtered.forEach(student => {
        const div = document.createElement('div');
        div.className = `student-card ${student.done ? 'done' : ''}`;
        
        div.innerHTML = `
            <div class="student-info">
                <h4>${student.name}</h4>
                <span>${student.id}</span>
            </div>
            <button class="capture-btn" onclick="startCapture('${student.id}')">
                ${student.done ? '✅ مكتمل' : '📸 تصوير'}
            </button>
        `;
        studentsList.appendChild(div);
    });

    // إظهار زر التحميل إذا كان هناك طلاب مكتملين
    if (students.some(s => s.done)) {
        downloadZipBtn.style.display = 'block';
    }
}

searchInput.addEventListener('input', (e) => renderStudentsList(e.target.value));

// 3. تهيئة الذكاء الاصطناعي (MediaPipe Face Mesh)
async function initFaceMesh() {
    faceMesh = new FaceMesh({locateFile: (file) => {
        return `https://cdn.jsdelivr.net/npm/@mediapipe/face_mesh/${file}`;
    }});
    
    faceMesh.setOptions({
        maxNumFaces: 1,
        refineLandmarks: true,
        minDetectionConfidence: 0.7,
        minTrackingConfidence: 0.7
    });
    
    faceMesh.onResults(onResults);
}

// 4. بدء جلسة التصوير للطالب
function startCapture(studentId) {
    currentStudent = students.find(s => s.id === studentId);
    if (!currentStudent) return;
    
    currentStudentName.textContent = currentStudent.name;
    stepStudents.classList.remove('active');
    stepCamera.classList.add('active');
    
    // تصفير المتغيرات
    currentPoseIndex = 0;
    captureHoldFrames = 0;
    centerYaw = 1.0;
    centerPitch = 1.0;
    resetPoseBadges();
    updateInstruction();
    
    // تشغيل الكاميرا الخلفية
    startCamera();
}

backToListBtn.addEventListener('click', () => {
    stopCamera();
    stepCamera.classList.remove('active');
    stepStudents.classList.add('active');
    renderStudentsList(searchInput.value);
});

function startCamera() {
    videoElement.classList.add('environment');
    canvasElement.classList.add('environment');

    camera = new Camera(videoElement, {
        onFrame: async () => {
            if (faceMesh) {
                await faceMesh.send({image: videoElement});
            }
        },
        width: 1280,
        height: 720,
        facingMode: 'environment' // كاميرا خلفية
    });
    camera.start();
}

function stopCamera() {
    if (camera) {
        camera.stop();
        camera = null;
    }
}

function distance(p1, p2) {
    return Math.sqrt(Math.pow(p1.x - p2.x, 2) + Math.pow(p1.y - p2.y, 2));
}

// 5. معالجة الإطارات واستخراج الزوايا
function onResults(results) {
    canvasElement.width = videoElement.videoWidth;
    canvasElement.height = videoElement.videoHeight;
    canvasCtx.save();
    canvasCtx.clearRect(0, 0, canvasElement.width, canvasElement.height);
    canvasCtx.drawImage(results.image, 0, 0, canvasElement.width, canvasElement.height);
    
    if (results.multiFaceLandmarks && results.multiFaceLandmarks.length > 0) {
        const landmarks = results.multiFaceLandmarks[0];
        
        // النقاط الأساسية
        const nose = landmarks[1];
        const leftEye = landmarks[33]; // العين يمين الصورة
        const rightEye = landmarks[263]; // العين يسار الصورة
        const chin = landmarks[152];
        const eyeMid = { x: (leftEye.x + rightEye.x) / 2, y: (leftEye.y + rightEye.y) / 2 };
        
        // حساب النسب
        const distLeft = distance(nose, leftEye);
        const distRight = distance(nose, rightEye);
        const yaw = distLeft / (distRight + 0.0001);
        
        const distUp = distance(nose, eyeMid);
        const distDown = distance(nose, chin);
        const pitch = distUp / (distDown + 0.0001);
        
        // رسم النقاط (اختياري للزينة)
        canvasCtx.fillStyle = '#2ecc71';
        [nose, leftEye, rightEye, chin].forEach(p => {
            canvasCtx.beginPath();
            canvasCtx.arc(p.x * canvasElement.width, p.y * canvasElement.height, 4, 0, 2 * Math.PI);
            canvasCtx.fill();
        });

        checkPose(yaw, pitch);
        
    } else {
        guideCircle.classList.remove('success');
        captureHoldFrames = 0;
        captureProgress.style.width = '0%';
        if(currentPoseIndex < poses.length) {
            instructionText.textContent = "الرجاء توجيه الكاميرا نحو الوجه بوضوح";
            instructionText.style.color = "#f39c12";
        }
    }
    
    canvasCtx.restore();
}

function checkPose(yaw, pitch) {
    if (currentPoseIndex >= poses.length) return;
    
    const requiredPose = poses[currentPoseIndex];
    let isMatching = false;
    
    // معايرة الديناميكية (Dynamic Calibration)
    if (requiredPose === 'center') {
        if (yaw > 0.7 && yaw < 1.3 && pitch > 0.5 && pitch < 1.5) {
            isMatching = true;
            // تحديث المركز باستمرار أثناء التثبيت
            centerYaw = yaw;
            centerPitch = pitch;
        }
    } else {
        // حساب نسبة الميلان مقارنة بالمركز المخصص لهذا الطالب
        const relYaw = yaw / centerYaw;
        const relPitch = pitch / centerPitch;
        
        // بما أن الكاميرا خلفية (بدون مرآة):
        // تحريك الرأس لليمين (بالنسبة للناظر) يجعل الأنف يقترب من العين اليمنى (rightEye in mediapipe = 263)
        // Mediapipe coordinates: x=0 is left of image, x=1 is right of image.
        // rightEye (263) is on the right side of the face (from user perspective it's left of image if mirrored).
        // Since we are NOT mirrored:
        // Student turns right -> Nose moves towards left edge of image -> distLeft decreases, distRight increases -> relYaw decreases (< 0.6)
        // Student turns left -> Nose moves towards right edge of image -> distLeft increases, distRight decreases -> relYaw increases (> 1.6)
        
        if (requiredPose === 'right' && relYaw < 0.65) isMatching = true; // تعديل حسب الكاميرا الخلفية
        if (requiredPose === 'left' && relYaw > 1.45) isMatching = true;
        if (requiredPose === 'up' && relPitch < 0.65) isMatching = true;
        if (requiredPose === 'down' && relPitch > 1.35) isMatching = true;
    }
    
    if (isMatching) {
        guideCircle.classList.add('success');
        instructionText.style.color = "#2ecc71";
        instructionText.textContent = "ممتاز.. إثبت!";
        captureHoldFrames++;
        captureProgress.style.width = `${(captureHoldFrames / REQUIRED_HOLD_FRAMES) * 100}%`;
        
        if (captureHoldFrames >= REQUIRED_HOLD_FRAMES) {
            takeSnapshot(requiredPose);
        }
    } else {
        guideCircle.classList.remove('success');
        updateInstruction();
        captureHoldFrames = 0;
        captureProgress.style.width = '0%';
    }
}

function updateInstruction() {
    if (currentPoseIndex >= poses.length) return;
    instructionText.textContent = instructions[poses[currentPoseIndex]];
    instructionText.style.color = "white";
    
    // تحديث الشارات
    document.querySelectorAll('.pose-badge').forEach(b => {
        b.classList.remove('active');
    });
    document.getElementById(`pose-${poses[currentPoseIndex]}`).classList.add('active');
}

function resetPoseBadges() {
    document.querySelectorAll('.pose-badge').forEach(b => {
        b.className = 'pose-badge pending';
    });
}

function takeSnapshot(poseName) {
    captureHoldFrames = 0;
    captureProgress.style.width = '0%';
    
    // وميض الشاشة
    const flash = document.createElement('div');
    flash.style.position = 'absolute';
    flash.style.top = '0'; flash.style.left = '0';
    flash.style.width = '100%'; flash.style.height = '100%';
    flash.style.backgroundColor = 'white';
    flash.style.zIndex = '100';
    stepCamera.appendChild(flash);
    setTimeout(() => flash.remove(), 150);
    
    // التقاط الصورة
    const dataUrl = canvasElement.toDataURL('image/jpeg', 0.9);
    const base64Data = dataUrl.replace(/^data:image\/(png|jpeg);base64,/, "");
    
    // الاسم: 123456_1_center.jpg
    const idx = currentPoseIndex + 1;
    const filename = `${currentStudent.id}_${idx}_${poseName}.jpg`;
    
    // إضافة الصورة للملف المضغوط
    zip.file(filename, base64Data, {base64: true});
    capturedCount++;
    
    // تحديد الشارة كمكتملة
    document.getElementById(`pose-${poseName}`).classList.remove('active');
    document.getElementById(`pose-${poseName}`).classList.add('done');
    
    currentPoseIndex++;
    
    if (currentPoseIndex >= poses.length) {
        // انتهى التصوير لهذا الطالب
        instructionText.textContent = "✅ اكتمل التصوير بنجاح!";
        instructionText.style.color = "#2ecc71";
        
        currentStudent.done = true;
        
        setTimeout(() => {
            stopCamera();
            stepCamera.classList.remove('active');
            stepStudents.classList.add('active');
            renderStudentsList(searchInput.value);
        }, 1500);
    } else {
        updateInstruction();
    }
}

// 6. تصدير الملف المضغوط (ZIP)
downloadZipBtn.addEventListener('click', () => {
    if (capturedCount === 0) {
        alert("لم يتم التقاط أي صور بعد!");
        return;
    }
    
    downloadZipBtn.innerHTML = "⏳ جاري إنشاء الملف...";
    downloadZipBtn.style.pointerEvents = "none";
    
    zip.generateAsync({type:"blob"}).then(function(content) {
        // إنشاء رابط التحميل
        const link = document.createElement('a');
        link.href = URL.createObjectURL(content);
        link.download = `طلاب_مساعد_الحصر_${new Date().getTime()}.zip`;
        link.click();
        
        downloadZipBtn.innerHTML = "✅ تم التنزيل بنجاح!";
        setTimeout(() => {
            downloadZipBtn.innerHTML = "⬇️ تنزيل جميع الصور كملف (ZIP)";
            downloadZipBtn.style.pointerEvents = "auto";
        }, 3000);
    });
});
