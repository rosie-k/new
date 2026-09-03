#!/usr/bin/env bash

# CLI Login Helper Script
# This script guides you through authenticating with GitHub CLI (gh) and Vercel CLI (vercel).

# Set colors for styling
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[0;33m'
RED='\033[0;31m'
BOLD='\033[1m'
NC='\033[0m' # No Color

echo -e "${BLUE}${BOLD}====================================================${NC}"
echo -e "${BLUE}${BOLD}      CLI Authentication Helper (GitHub & Vercel)    ${NC}"
echo -e "${BLUE}${BOLD}====================================================${NC}"
echo ""

# Function to handle GitHub CLI Login
login_github() {
    echo -e "${GREEN}${BOLD}[1/2] GitHub CLI (gh) Login${NC}"
    echo -e "Choose how you want to log in to GitHub CLI:"
    echo -e "1) Interactive Web-based Login (Recommended)"
    echo -e "2) Log in using a Personal Access Token (PAT)"
    echo -e "3) Check current login status"
    echo -e "4) Skip"
    read -p "Enter your choice (1-4): " gh_choice

    case $gh_choice in
        1)
            echo -e "\n${YELLOW}Starting interactive login. Follow the terminal prompts...${NC}"
            gh auth login
            ;;
        2)
            echo -e "\n${YELLOW}Please enter your GitHub Personal Access Token (PAT):${NC}"
            read -s -p "GitHub Token: " gh_token
            echo ""
            if [ -z "$gh_token" ]; then
                echo -e "${RED}Token cannot be empty.${NC}"
            else
                echo "$gh_token" | gh auth login --with-token
                if [ $? -eq 0 ]; then
                    echo -e "${GREEN}Successfully authenticated with GitHub CLI!${NC}"
                else
                    echo -e "${RED}Authentication failed. Please check your token.${NC}"
                fi
            fi
            ;;
        3)
            echo -e "\n${BLUE}Checking GitHub login status...${NC}"
            gh auth status
            ;;
        *)
            echo -e "\nSkipping GitHub CLI login."
            ;;
    esac
}

# Function to handle Vercel CLI Login
login_vercel() {
    echo -e "\n${GREEN}${BOLD}[2/2] Vercel CLI (vercel) Login${NC}"
    echo -e "Choose how you want to log in to Vercel CLI:"
    echo -e "1) Interactive Email-based Login (Recommended)"
    echo -e "2) Log in using a Vercel Token"
    echo -e "3) Check current login status"
    echo -e "4) Skip"
    read -p "Enter your choice (1-4): " vercel_choice

    case $vercel_choice in
        1)
            echo -e "\n${YELLOW}Starting interactive Vercel login. Enter your email when prompted...${NC}"
            vercel login
            ;;
        2)
            echo -e "\n${YELLOW}Please enter your Vercel Access Token (created at vercel.com/account/tokens):${NC}"
            read -s -p "Vercel Token: " vercel_token
            echo ""
            if [ -z "$vercel_token" ]; then
                echo -e "${RED}Token cannot be empty.${NC}"
            else
                echo -e "${YELLOW}Vercel uses token environment variable for headless login.${NC}"
                echo -e "To use this session, run: ${BOLD}export VERCEL_TOKEN=$vercel_token${NC}"
                echo -e "Verifying token validity..."
                export VERCEL_TOKEN=$vercel_token
                vercel whoami
            fi
            ;;
        3)
            echo -e "\n${BLUE}Checking Vercel login status...${NC}"
            vercel whoami
            ;;
        *)
            echo -e "\nSkipping Vercel CLI login."
            ;;
    esac
}

# Main Execution Flow
if ! command -v gh &> /dev/null; then
    echo -e "${RED}GitHub CLI (gh) is not installed or not in PATH.${NC}"
else
    login_github
fi

echo -e "\n----------------------------------------------------\n"

if ! command -v vercel &> /dev/null; then
    echo -e "${RED}Vercel CLI (vercel) is not installed or not in PATH.${NC}"
else
    login_vercel
fi

echo -e "\n${BLUE}${BOLD}====================================================${NC}"
echo -e "${GREEN}${BOLD}                   Process Finished!                 ${NC}"
echo -e "${BLUE}${BOLD}====================================================${NC}"
