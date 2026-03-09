# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs
import torch
import torchvision

# Get PyTorch paths
torch_path = os.path.dirname(torch.__file__)
torchvision_path = os.path.dirname(torchvision.__file__)

# Detect CUDA availability during build (optional, but helpful)
cuda_available = torch.cuda.is_available() if hasattr(torch, 'cuda') else False
print(f"CUDA available during build: {cuda_available}")

# Collect all necessary NVIDIA/CUDA binaries
def get_nvidia_binaries():
    nvidia_bins = []
    
    # Common CUDA library patterns
    cuda_libs = [
        'cudart64_*.dll',
        'cublas64_*.dll',
        'cublasLt64_*.dll',
        'cudnn64_*.dll',
        'cufft64_*.dll',
        'curand64_*.dll',
        'cusolver64_*.dll',
        'cusparse64_*.dll',
        'nvrtc64_*.dll',
        'nvrtc-builtins64_*.dll',
        'cuda*.dll',
    ]
    
    # Search in common locations
    search_paths = [
        os.path.join(torch_path, 'lib'),
        os.path.join(torch_path, '..', 'nvidia', 'cuda_runtime', 'bin'),
        os.path.join(torch_path, '..', 'nvidia', 'cublas', 'bin'),
        os.path.join(torch_path, '..', 'nvidia', 'cudnn', 'bin'),
        os.path.join(torch_path, '..', 'nvidia', 'cufft', 'bin'),
        os.path.join(torch_path, '..', 'nvidia', 'curand', 'bin'),
        os.path.join(torch_path, '..', 'nvidia', 'cusolver', 'bin'),
        os.path.join(torch_path, '..', 'nvidia', 'cusparse', 'bin'),
        os.path.join(torch_path, '..', 'nvidia', 'nvrtc', 'bin'),
        os.environ.get('CUDA_PATH', ''),
        os.path.join(os.environ.get('CUDA_PATH', ''), 'bin'),
        'C:\\Program Files\\NVIDIA GPU Computing Toolkit\\CUDA\\v11.8\\bin',
        'C:\\Program Files\\NVIDIA GPU Computing Toolkit\\CUDA\\v12.1\\bin',
    ]
    
    for path in search_paths:
        if os.path.exists(path):
            for pattern in cuda_libs:
                import glob
                found = glob.glob(os.path.join(path, pattern))
                nvidia_bins.extend(found)
    
    return nvidia_bins

# PyInstaller configuration
a = Analysis(
    ['egg_counter.py'],
    pathex=['.'],
    binaries=[
        # Include NVIDIA/CUDA DLLs
        *[(lib, 'nvidia') for lib in get_nvidia_binaries()],
        # Include OpenCV binaries
        *collect_dynamic_libs('cv2'),
        *collect_dynamic_libs('ultralytics'),
    ],
    datas=[
        # Include any necessary data files
        *collect_data_files('torch', include_py_files=True),
        *collect_data_files('torchvision', include_py_files=True),
        *collect_data_files('ultralytics', include_py_files=True),
    ],
    hiddenimports=[
        # Core imports
        'ultralytics',
        'ultralytics.nn.tasks',
        'ultralytics.yolo',
        'ultralytics.yolo.engine',
        'ultralytics.yolo.engine.model',
        'ultralytics.yolo.engine.predictor',
        'ultralytics.yolo.engine.trainer',
        'ultralytics.yolo.engine.validator',
        'ultralytics.yolo.engine.results',
        'ultralytics.yolo.utils',
        'ultralytics.yolo.utils.ops',
        'ultralytics.yolo.utils.torch_utils',
        'ultralytics.yolo.utils.loss',
        'ultralytics.yolo.utils.metrics',
        'ultralytics.yolo.utils.plotting',
        'ultralytics.yolo.utils.callbacks',
        'ultralytics.yolo.utils.files',
        'ultralytics.yolo.utils.downloads',
        
        # PyTorch and related
        'torch',
        'torchvision',
        'torchvision.ops',
        'torch.nn',
        'torch.nn.modules',
        'torch.jit',
        'torch._C',
        'torch._C._nn',
        'torch._C._fft',
        'torch._C._linalg',
        'torch._C._sparse',
        'torch._C._cuda',
        'torch._C._nvrtc',
        
        # NumPy
        'numpy',
        'numpy.core._multiarray_umath',
        'numpy.core._multiarray_tests',
        'numpy.core._dtype_ctypes',
        'numpy.core._methods',
        'numpy.core.umath_tests',
        
        # OpenCV
        'cv2',
        'cv2.gapi',
        
        # Harvesters (for Delta camera)
        'harvesters',
        'harvesters.core',
        'harvesters._gentl',
        'harvesters._event',
        'harvesters._buffer',
        'harvesters._device',
        'harvesters._system',
        'harvesters._image',
        
        # PyModbus
        'pyModbusTCP',
        'pyModbusTCP.client',
        'pyModbusTCP.constants',
        'pyModbusTCP.utils',
        
        # Tkinter
        'PIL',
        'PIL._tkinter_finder',
        'PIL.Image',
        'PIL.ImageTk',
        'PIL._tkinter',
        
        # Other dependencies
        'json',
        'csv',
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.styles',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'openpyxl.reader.excel',
        'openpyxl.reader.workbook',
        'openpyxl.writer.excel',
        'openpyxl.writer.workbook',
        'collections',
        'collections.abc',
        'queue',
        'threading',
        'datetime',
        'time',
        'argparse',
        'glob',
        'os',
        'sys',
        'socketserver',
        'http',
        'http.server',
        'logging',
        'subprocess',
        'ctypes',
        'ctypes.wintypes',
        'deque',
        'typing',
        'dataclasses',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

# Add PyTorch CUDA libraries explicitly
for root, dirs, files in os.walk(os.path.join(torch_path, 'lib')):
    for file in files:
        if file.endswith('.dll') and any(x in file.lower() for x in ['cuda', 'cudnn', 'cublas', 'cufft', 'curand', 'cusolver', 'nvrtc']):
            a.binaries.append((os.path.join('torch', 'lib', file), os.path.join(root, file), 'BINARY'))

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='DAC_Egg_Counter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Set to False if you want to hide console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None  # Add path to your .ico file if you have one
)
