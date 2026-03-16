const bcrypt = require('bcrypt');
const { run } = require('./database');

async function seedAdvisors() {
    try {
        const defaultPassword = 'advisor123';
        const hashedPassword = await bcrypt.hash(defaultPassword, 10);

        await run(`INSERT INTO advisors (name, email, password, is_active, created_at) VALUES (?, ?, ?, 1, ?)`,
            ['Ravi Kant', 'ravi@incremintedge.com', hashedPassword, new Date().toISOString()]);

        console.log('✅ Default advisor created successfully!');
        console.log('📧 Email: ravi@incremintedge.com');
        console.log('🔑 Password: advisor123');
        console.log('⚠️  Please change password after first login');
        
        process.exit(0);
    } catch (error) {
        if (error.message && error.message.includes('UNIQUE')) {
            console.log('ℹ️  Default advisor already exists');
        } else {
            console.error('❌ Error creating advisor:', error.message);
        }
        process.exit(1);
    }
}

seedAdvisors();
