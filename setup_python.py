#!/usr/bin/env python3
"""
Setup script to fix python-pptx compatibility with Python 3.12
"""

import sys
import subprocess
import collections.abc

def patch_collections_compatibility():
    """
    Patch the collections compatibility issue for python-pptx
    """
    try:
        import pptx.compat
        
        # Check if the patch is needed
        if not hasattr(collections, 'abc'):
            print("⚠️ Python 3.12+ detected - patching collections compatibility...")
            
            # Monkey patch for compatibility
            import collections
            collections.Container = collections.abc.Container
            collections.Iterable = collections.abc.Iterable
            collections.Mapping = collections.abc.Mapping
            collections.MutableMapping = collections.abc.MutableMapping
            collections.Sequence = collections.abc.Sequence
            
            print("✅ Collections compatibility patched")
        else:
            print("✅ Collections compatibility already available")
            
    except ImportError as e:
        print(f"❌ python-pptx not installed: {e}")
        return False
    except Exception as e:
        print(f"❌ Error patching collections: {e}")
        return False
    
    return True

def install_dependencies():
    """
    Install Python dependencies with correct versions
    """
    print("📦 Installing Python dependencies...")
    
    packages = [
        "python-pptx>=0.6.22",
        "Pillow>=10.0.1"
    ]
    
    for package in packages:
        try:
            print(f"Installing {package}...")
            result = subprocess.run([
                sys.executable, "-m", "pip", "install", "--upgrade", package
            ], capture_output=True, text=True, check=True)
            print(f"✅ {package} installed successfully")
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to install {package}: {e}")
            print(f"stdout: {e.stdout}")
            print(f"stderr: {e.stderr}")
            return False
    
    return True

def test_imports():
    """
    Test if all required libraries can be imported
    """
    print("\n🧪 Testing imports...")
    
    try:
        from pptx import Presentation
        print("✅ python-pptx imported successfully")
    except Exception as e:
        print(f"❌ Failed to import python-pptx: {e}")
        return False
    
    try:
        from PIL import Image
        print("✅ Pillow imported successfully")
    except Exception as e:
        print(f"❌ Failed to import Pillow: {e}")
        return False
    
    return True

if __name__ == '__main__':
    print("🐍 Python Environment Setup for TCNG Proposal Generator")
    print(f"Python version: {sys.version}")
    
    # Install dependencies
    if not install_dependencies():
        print("❌ Failed to install dependencies")
        sys.exit(1)
    
    # Apply compatibility patches
    if not patch_collections_compatibility():
        print("❌ Failed to patch compatibility")
        sys.exit(1)
    
    # Test imports
    if not test_imports():
        print("❌ Import test failed")
        sys.exit(1)
    
    print("\n✅ Python environment setup complete!")
    print("🎉 Ready to generate proposals!")