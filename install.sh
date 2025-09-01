#!/bin/bash

echo "🚀 Installing TCNG Proposal Generator Dependencies..."

# Install Node.js dependencies
echo "📦 Installing Node.js dependencies..."
npm install

# Install Python dependencies
echo "📦 Installing Python dependencies..."
pip install --upgrade pip

# Try to install the latest compatible version
pip install --upgrade "python-pptx>=0.6.22" "Pillow>=10.0.1"

# Alternative: Install specific working versions
if [ $? -ne 0 ]; then
    echo "⚠️ Trying alternative python-pptx version..."
    pip install --upgrade python-pptx==0.6.23 Pillow==10.0.1
fi

# Test the installation
echo "🧪 Testing Python imports..."
python3 -c "
try:
    import collections.abc
    import collections
    if not hasattr(collections, 'Container'):
        collections.Container = collections.abc.Container
        collections.Iterable = collections.abc.Iterable  
        collections.Mapping = collections.abc.Mapping
        collections.MutableMapping = collections.abc.MutableMapping
        collections.Sequence = collections.abc.Sequence
        print('✅ Collections compatibility patched')
    
    from pptx import Presentation
    from PIL import Image
    print('✅ All Python dependencies imported successfully!')
    
except Exception as e:
    print(f'❌ Import error: {e}')
    exit(1)
"

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ Installation completed successfully!"
    echo ""
    echo "🎯 To start the server:"
    echo "   npm start"
    echo ""
    echo "📋 Make sure you have TP_Template.pptx in the current directory"
    echo "   with a {{TP_MSB}} placeholder for the image"
else
    echo "❌ Installation failed. Please check the error messages above."
    exit 1
fi