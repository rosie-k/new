#!/usr/bin/env bash

# Elegant Release & Deployment Automation Script
# Connects your local project with GitHub (via gh) and Vercel (via vercel).

# Styling colors
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[0;33m'
RED='\033[0;31m'
BOLD='\033[1m'
NC='\033[0m' # No Color

clear
echo -e "${BLUE}${BOLD}====================================================${NC}"
echo -e "${BLUE}${BOLD}        ELEGANT RELEASER (GitHub & Vercel)          ${NC}"
echo -e "${BLUE}${BOLD}====================================================${NC}"
echo ""

# 1. Verification of installed tools
echo -e "${BLUE}[Step 1/5] Checking Required Tools...${NC}"

if ! command -v gh &> /dev/null; then
    echo -e "${RED}❌ GitHub CLI (gh) is not installed. Please install it first.${NC}"
    exit 1
fi

if ! command -v vercel &> /dev/null; then
    echo -e "${RED}❌ Vercel CLI (vercel) is not installed. Please install it first.${NC}"
    exit 1
fi

echo -e "${GREEN}✅ GitHub CLI (gh) is available.${NC}"
echo -e "${GREEN}✅ Vercel CLI (vercel) is available.${NC}"
echo ""

# 2. Check Authentication
echo -e "${BLUE}[Step 2/5] Checking Authentication Status...${NC}"

# Check GitHub Status
gh auth status &> /dev/null
GH_AUTH=$?
if [ $GH_AUTH -ne 0 ]; then
    echo -e "${YELLOW}⚠️  GitHub CLI is not authenticated.${NC}"
    echo -e "Please run ${BOLD}./cli_login.sh${NC} first to log in, or authenticate now by entering: ${BOLD}gh auth login${NC}"
    read -p "Would you like to run 'gh auth login' now? (y/n): " gh_login_now
    if [[ $gh_login_now =~ ^[Yy]$ ]]; then
        gh auth login
    else
        echo -e "${RED}Aborting. GitHub authentication is required to release.${NC}"
        exit 1
    fi
else
    echo -e "${GREEN}✅ GitHub: Authenticated successfully.${NC}"
fi

# Check Vercel Status
vercel whoami &> /dev/null
VERCEL_AUTH=$?
if [ $VERCEL_AUTH -ne 0 ] && [ -z "$VERCEL_TOKEN" ]; then
    echo -e "${YELLOW}⚠️  Vercel CLI is not authenticated.${NC}"
    echo -e "Please run ${BOLD}./cli_login.sh${NC} first to log in, or authenticate now by entering: ${BOLD}vercel login${NC}"
    read -p "Would you like to run 'vercel login' now? (y/n): " vercel_login_now
    if [[ $vercel_login_now =~ ^[Yy]$ ]]; then
        vercel login
    else
        echo -e "${RED}Aborting. Vercel authentication is required to deploy.${NC}"
        exit 1
    fi
else
    echo -e "${GREEN}✅ Vercel: Authenticated successfully.${NC}"
fi
echo ""

# 3. Setup Git Repository
echo -e "${BLUE}[Step 3/5] Checking Local Git Repository...${NC}"
if [ ! -d .git ]; then
    echo -e "${YELLOW}Initializing new Git repository...${NC}"
    git init
    git branch -M main
fi

# Add all changes if dirty
if [[ -n $(git status --porcelain) ]]; then
    echo -e "${YELLOW}Staging files and creating initial commit...${NC}"
    git add .
    read -p "Enter commit message [Default: 'Release: Initial deployment']: " commit_msg
    if [ -z "$commit_msg" ]; then
         commit_msg="Release: Initial deployment"
    fi
    git commit -m "$commit_msg"
    echo -e "${GREEN}✅ Files committed successfully.${NC}"
else
    echo -e "${GREEN}✅ Git workspace is clean.${NC}"
fi
echo ""

# 4. GitHub Release
echo -e "${BLUE}[Step 4/5] Publishing to GitHub...${NC}"
# Check if remote exists
REMOTE_EXISTS=$(git remote | grep origin)
if [ -z "$REMOTE_EXISTS" ]; then
    echo -e "No GitHub remote 'origin' found."
    read -p "Would you like to create a NEW repository on GitHub for this project? (y/n): " create_repo
    if [[ $create_repo =~ ^[Yy]$ ]]; then
        read -p "Enter repository name [Default: 'gongyoonuri-booking-dashboard']: " repo_name
        if [ -z "$repo_name" ]; then
            repo_name="gongyoonuri-booking-dashboard"
        fi
        echo -e "${YELLOW}Creating repository '$repo_name' on GitHub...${NC}"
        gh repo create "$repo_name" --public --source=. --remote=origin --push
        echo -e "${GREEN}✅ GitHub repository created and pushed!${NC}"
    else
        echo -e "Skipping GitHub repository creation. If you already have one, run: ${BOLD}git remote add origin <URL>${NC}"
    fi
else
    echo -e "${GREEN}Using existing Git remote 'origin'.${NC}"
    echo -e "${YELLOW}Pushing commits to remote main branch...${NC}"
    git push -u origin main
fi

# Create high-level release/tag on GitHub
read -p "Would you like to tag and create a new Release on GitHub? (y/n): " make_release
if [[ $make_release =~ ^[Yy]$ ]]; then
    read -p "Enter release version tag [Default: 'v1.0.0']: " version_tag
    if [ -z "$version_tag" ]; then
        version_tag="v1.0.0"
    fi
    read -p "Enter release title [Default: 'Release v1.0.0']: " release_title
    if [ -z "$release_title" ]; then
        release_title="Release $version_tag"
    fi
    echo -e "${YELLOW}Creating GitHub Release '$version_tag'...${NC}"
    gh release create "$version_tag" --title "$release_title" --notes "Release generated automatically via Elegant Releaser Script"
    echo -e "${GREEN}✅ GitHub Release created successfully!${NC}"
fi
echo ""

# 5. Vercel Deploy
echo -e "${BLUE}[Step 5/5] Deploying to Vercel Production...${NC}"
read -p "Would you like to deploy this app to Vercel now? (y/n): " deploy_vercel
if [[ $deploy_vercel =~ ^[Yy]$ ]]; then
    echo -e "${YELLOW}Deploying to Vercel Production... (This may take a minute)${NC}"
    vercel --prod --yes
    echo -e "${GREEN}✅ Vercel Deployment Completed!${NC}"
else
    echo -e "Vercel deployment skipped."
fi

echo ""
echo -e "${BLUE}${BOLD}====================================================${NC}"
echo -e "${GREEN}${BOLD}       🎉 All Specified Release Flows Finished!     ${NC}"
echo -e "${BLUE}${BOLD}====================================================${NC}"
