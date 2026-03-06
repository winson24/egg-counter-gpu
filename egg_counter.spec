# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

# Collect PyTorch CUDA dependencies
torch_libs = collect_dynamic_libs('torch')
torch_data = collect_data_files('torch')
cuda_libs = []

# Add CUDA libraries if they exist
cuda_paths = [
    os.environ.get('CUDA_PATH', ''),
    'C:/Program Files/NVIDIA GPU Computing Toolkit/CUDA/v11.8',
    'C:/Program Files/NVIDIA GPU Computing Toolkit/CUDA/v12.1',
]

for cuda_path in cuda_paths:
    if cuda_path and os.path.exists(cuda_path):
        bin_path = os.path.join(cuda_path, 'bin')
        if os.path.exists(bin_path):
            for dll in os.listdir(bin_path):
                if dll.endswith('.dll') and any(x in dll.lower() for x in ['cuda', 'cudnn', 'cublas', 'cufft', 'curand', 'cusolver', 'cusp']):
                    cuda_libs.append((bin_path, [dll]))

a = Analysis(
    ['egg_counter.py'],
    pathex=[],
    binaries=[
        # Add PyTorch CUDA libraries
        *torch_libs,
        *cuda_libs,
    ],
    datas=[
        # Add any data files your model needs
        *torch_data,
    ],
    hiddenimports=[
        'torch',
        'torchvision',
        'torch.cuda',
        'torch.cuda.amp',
        'torch.nn',
        'torch.nn.parallel',
        'torch.jit',
        'torch.tensor',
        'torch.autograd',
        'torch.backends.cuda',
        'torch.backends.cudnn',
        'torch.distributed',
        'torch.distributions',
        'torch.fft',
        'torch.futures',
        'torch.linalg',
        'torch.mps',
        'torch.nn.intrinsic',
        'torch.nn.quantized',
        'torch.nn.utils',
        'torch.optim',
        'torch.signal',
        'torch.sparse',
        'torch.special',
        'torch.utils',
        'torch.utils.data',
        'torch.utils.cpp_extension',
        'torch.utils.data.datapipes',
        'torch.utils.hooks',
        'torch.utils.model_zoo',
        'torch.utils.tensorboard',
        'torch.onnx',
        'torch.overrides',
        'torch.package',
        'torch.profiler',
        'torch.testing',
        'torch.version',
        'ultralytics',
        'ultralytics.nn',
        'ultralytics.yolo',
        'ultralytics.yolo.v8',
        'ultralytics.yolo.v8.detect',
        'ultralytics.yolo.utils',
        'ultralytics.yolo.utils.torch_utils',
        'cv2',
        'cv2.cv2',
        'numpy',
        'numpy.core',
        'numpy.lib',
        'numpy.fft',
        'numpy.linalg',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'PIL._imagingtk',
        'PIL._tkinter_finder',
        'pyModbusTCP',
        'pyModbusTCP.client',
        'harvesters',
        'harvesters.core',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.scrolledtext',
        'collections',
        'collections.abc',
        'queue',
        'csv',
        'json',
        'openpyxl',
        'openpyxl.styles',
        'datetime',
        'argparse',
        'glob',
        'time',
        'threading',
        'socketserver',
        'http.server',
        'http',
        'logging',
        'subprocess',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

# Add NVIDIA DLLs specifically
for dll in [
    'cublas64_*.dll',
    'cublasLt64_*.dll',
    'cudart64_*.dll',
    'cudnn64_*.dll',
    'cudnn_adv_infer64_*.dll',
    'cudnn_adv_train64_*.dll',
    'cudnn_cnn_infer64_*.dll',
    'cudnn_cnn_train64_*.dll',
    'cudnn_ops_infer64_*.dll',
    'cudnn_ops_train64_*.dll',
    'cufft64_*.dll',
    'curand64_*.dll',
    'cusolver64_*.dll',
    'cusparse64_*.dll',
    'nvrtc64_*.dll',
    'nvrtc-builtins64_*.dll',
]:
    a.binaries += [('', dll, 'BINARY')]

pyz = PYZ(a.pure)

# Create the executable in a directory structure
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DAC_Egg_Counter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'  # Add your icon if you have one
)

# Create the one-directory distribution
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DAC_Egg_Counter'

)
