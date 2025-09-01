const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const { spawn } = require('child_process');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.use(express.static('.'));

// Create necessary directories
const createDirectories = async () => {
    const dirs = ['temp_images', 'generated_proposals', 'uploads'];
    for (const dir of dirs) {
        try {
            await fs.mkdir(dir, { recursive: true });
        } catch (error) {
            console.error(`Error creating directory ${dir}:`, error);
        }
    }
};

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: async (req, file, cb) => {
        await createDirectories();
        cb(null, 'temp_images/');
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
    }
});

const upload = multer({ 
    storage: storage,
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    }
});

// Serve the main HTML file
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Generate proposal endpoint
app.post('/api/generate-proposal', upload.single('image'), async (req, res) => {
    try {
        console.log('📝 Processing proposal generation request...');
        
        // Get form data from request body
        const {
            building_name = '',
            address = '',
            survey_date = '',
            prepared_by = '',
            prepared_date = '',
            type_building = '',
            building_manager_name = '',
            building_manager_email = '',
            building_manager_phone = '',
            building_manager_company = '',
            otic = '',
            tap_new_or_spare = '',
            tapping_location = '',
            tapping_location_level = '',
            site_assessment_mccb = '',
            tnb_meter = '',
            tnb_na = '',
            parking_location = '',
            ev_charger_model = '',
            network_strength = '',
            imageData = '' // Base64 image data from canvas
        } = req.body;

        console.log('📋 Form data received:', { building_name, address: address.substring(0, 50) });
        
        // Validate required template file
        const templatePath = path.join(__dirname, 'TP_Template.pptx');
        try {
            await fs.access(templatePath);
            console.log('✅ Template file found:', templatePath);
        } catch (error) {
            console.error('❌ Template file not found:', templatePath);
            return res.status(400).json({
                success: false,
                message: 'Template file TP_Template.pptx not found. Please ensure it exists in the project directory.'
            });
        }

        // Prepare data for Python script
        const proposalData = {
            building_name,
            address,
            survey_date,
            prepared_by,
            prepared_date,
            type_building,
            building_manager_name,
            building_manager_email,
            building_manager_phone,
            building_manager_company,
            otic,
            tap_new_or_spare,
            tapping_location,
            tapping_location_level,
            site_assessment_mccb,
            tnb_meter,
            tnb_na,
            parking_location,
            noofchargers: '2', // Fixed to 2 chargers
            ev_charger_model,
            network_strength
        };

        // Save image data if provided
        let imagePath = null;
        if (imageData && imageData.startsWith('data:image/')) {
            try {
                console.log('🖼️ Processing image data...');
                console.log('📊 Image data length:', imageData.length);
                console.log('📊 Image data prefix:', imageData.substring(0, 50));
                
                // Extract base64 data
                const base64Data = imageData.split(',')[1];
                const imageBuffer = Buffer.from(base64Data, 'base64');
                console.log('📊 Image buffer size:', imageBuffer.length, 'bytes');
                
                // Save to temp file
                const timestamp = Date.now();
                imagePath = path.join(__dirname, 'temp_images', `proposal_image_${timestamp}.png`);
                await fs.writeFile(imagePath, imageBuffer);
                console.log('✅ Image saved to:', imagePath);
                
                // Verify file was saved
                const stats = await fs.stat(imagePath);
                console.log('📁 Saved image file size:', stats.size, 'bytes');
                
            } catch (error) {
                console.error('❌ Error processing image:', error);
                return res.status(400).json({
                    success: false,
                    message: 'Error processing image data'
                });
            }
        } else {
            console.log('⚠️ No valid image data provided');
        }

        // Generate unique output filename
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const safeClientName = building_name.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20) || 'proposal';
        const outputFilename = `${safeClientName}_${timestamp}.pptx`;
        const outputPath = path.join(__dirname, 'generated_proposals', outputFilename);

        // Call Python script to generate PowerPoint
        console.log('🐍 Calling Python script to generate proposal...');
        
        const pythonArgs = [
            'proposal_processor.py',
            '--template', templatePath,
            '--output', outputPath,
            '--data', JSON.stringify(proposalData)
        ];

        if (imagePath) {
            pythonArgs.push('--image', imagePath);
        }

        const pythonProcess = spawn('python3', pythonArgs, {
            cwd: __dirname,
            stdio: ['pipe', 'pipe', 'pipe']
        });

        let stdout = '';
        let stderr = '';

        pythonProcess.stdout.on('data', (data) => {
            stdout += data.toString();
        });

        pythonProcess.stderr.on('data', (data) => {
            stderr += data.toString();
        });

        pythonProcess.on('close', async (code) => {
            console.log(`🐍 Python process completed with code: ${code}`);
            console.log('📝 Python stdout output:');
            console.log(stdout);
            if (stderr) {
                console.log('⚠️ Python stderr output:');
                console.log(stderr);
            }
            
            try {
                // Clean up temporary image file
                if (imagePath) {
                    try {
                        await fs.unlink(imagePath);
                        console.log('🗑️ Cleaned up temporary image file');
                    } catch (cleanupError) {
                        console.warn('⚠️ Could not clean up temporary image:', cleanupError);
                    }
                }

                if (code === 0) {
                    console.log('✅ Python script completed successfully');

                    // Check if output file was created
                    try {
                        await fs.access(outputPath);
                        const fileStats = await fs.stat(outputPath);
                        console.log('📁 Output file created successfully');
                        console.log('📊 Output file size:', fileStats.size, 'bytes');
                        
                        // Return file download response
                        res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"`);
                        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
                        
                        const fileBuffer = await fs.readFile(outputPath);
                        console.log('📤 Sending file to client');
                        res.send(fileBuffer);
                        
                        // Clean up output file after sending
                        setTimeout(async () => {
                            try {
                                await fs.unlink(outputPath);
                                console.log('🗑️ Cleaned up output file');
                            } catch (cleanupError) {
                                console.warn('⚠️ Could not clean up output file:', cleanupError);
                            }
                        }, 5000); // 5 second delay to ensure download completes

                    } catch (error) {
                        console.error('❌ Output file not found or not accessible:', outputPath);
                        console.error('File access error:', error);
                        res.status(500).json({
                            success: false,
                            message: 'Proposal generation completed but output file not found'
                        });
                    }
                } else {
                    console.error('❌ Python script failed with code:', code);
                    res.status(500).json({
                        success: false,
                        message: `Proposal generation failed (code ${code}): ${stderr || 'Unknown error'}`
                    });
                }
            } catch (error) {
                console.error('❌ Error in Python process completion handler:', error);
                res.status(500).json({
                    success: false,
                    message: 'Internal server error during proposal generation'
                });
            }
        });

        pythonProcess.on('error', (error) => {
            console.error('❌ Python process error:', error);
            res.status(500).json({
                success: false,
                message: 'Failed to start Python script. Please ensure Python 3 is installed.'
            });
        });

    } catch (error) {
        console.error('❌ Unexpected error in proposal generation:', error);
        res.status(500).json({
            success: false,
            message: 'Unexpected server error'
        });
    }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'ok', 
        timestamp: new Date().toISOString(),
        message: 'TCNG Proposal Generator API is running'
    });
});

// Error handling middleware
app.use((error, req, res, next) => {
    if (error instanceof multer.MulterError) {
        if (error.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({
                success: false,
                message: 'File too large. Maximum size is 10MB.'
            });
        }
    }
    
    console.error('Server error:', error);
    res.status(500).json({
        success: false,
        message: 'Internal server error'
    });
});

// Initialize directories and start server
createDirectories().then(() => {
    app.listen(PORT, () => {
        console.log('🚀 TCNG Proposal Generator Server Started');
        console.log(`📍 Server running at: http://localhost:${PORT}`);
        console.log(`📋 API endpoint: http://localhost:${PORT}/api/generate-proposal`);
        console.log(`💚 Health check: http://localhost:${PORT}/api/health`);
        console.log('🎯 Ready to generate proposals!');
        console.log('='*60);
    });
}).catch(error => {
    console.error('❌ Failed to initialize server:', error);
    process.exit(1);
});